import {bootstrap} from 'global-agent';
import {AxiosError, AxiosInstance, AxiosRequestConfig} from "axios"
import {HttpsProxyAgent} from "https-proxy-agent";
import {globalAgent} from "node:https"
import debug from "debug"
import { TLSSocket } from "node:tls";
const log = debug("genOffice").extend("requestHelper")
global.GLOBAL_AGENT_FORCE_GLOBAL_AGENT=false
bootstrap();

// Helper to handle circular references when logging
function getCircularReplacer() {
    const seen = new WeakSet();
    return (key : any, value : any) => {
        if (typeof value === "object" && value !== null) {
            if (seen.has(value)) {
                return "[Circular]";
            }
            seen.add(value);
        }
        return value;
    };
}

//Cleans objects for use in log entries
function clean(object :any) {
    if(object) {
        return JSON.parse(JSON.stringify(object, getCircularReplacer()));
    }
    return object;
}

//Attempts to get proxy settings from standard environment variables
export function getProxyURL() {
    return process.env.HTTPS_PROXY || process.env.https_proxy || process.env.HTTP_PROXY || process.env.http_proxy || 
           (process.env['https.proxyHost'] && process.env['https.proxyPort'] ? 
            `http://${process.env['https.proxyHost']}:${process.env['https.proxyPort']}` : undefined);
}

//If proxy is available, set global-agent flags and add agent
export function addProxy(config: AxiosRequestConfig) {
    const httpsProxy = getProxyURL();
    if(hasProxy()) {
        log("Adding Proxy %s",httpsProxy)
        global.GLOBAL_AGENT_FORCE_GLOBAL_AGENT=true;
        if(global.GLOBAL_AGENT) {
            global.GLOBAL_AGENT.HTTP_PROXY=httpsProxy
        }
        config.httpsAgent = new HttpsProxyAgent(httpsProxy!,{keepAlive:false});
    }
    addLogging(config);
}

//Breaks down tls information into smaller chunks for log analysis.
function debugTlsSocket(tlsSocket : TLSSocket) {
    return {
        // Basic connection info
        authorized: tlsSocket.authorized,
        authorizationError: tlsSocket.authorizationError,
        protocol: tlsSocket.getProtocol(),
        cipher: clean(tlsSocket.getCipher()),
        certificate: clean(tlsSocket.getPeerCertificate(true)), // verbose=true
        localCertificate: clean(tlsSocket.getCertificate()),

        // Connection details
        remoteAddress: tlsSocket.remoteAddress,
        remotePort: tlsSocket.remotePort,
        localAddress: tlsSocket.localAddress,
        localPort: tlsSocket.localPort,

        // SSL/TLS session info
        sessionId: tlsSocket.getSession()?.toString('hex'),
        tlsRenegotiated: tlsSocket.isSessionReused(),

        // Socket state
        connecting: tlsSocket.connecting,
        destroyed: tlsSocket.destroyed,
        readable: tlsSocket.readable,
        writable: tlsSocket.writable,

        // ALPN/NPN protocols
        alpnProtocol: tlsSocket.alpnProtocol,

        // Timeout settings
        timeout: tlsSocket.timeout,
        symbols : Object.getOwnPropertySymbols(tlsSocket)
            .filter(symbol => { return typeof symbol !== "function"})
            .map(symbol => {
                return {name:symbol.toString(),value:JSON.parse(JSON.stringify(tlsSocket[symbol], getCircularReplacer()))}
            })
    };
}

//Check to see if a proxy is available.
export function hasProxy() {
    const httpsProxy = getProxyURL();
    const result = !!httpsProxy;
    return result;
}

//Replace config https agent with global agent and remove global-proxy settings.
export function removeProxy(config: AxiosRequestConfig) {
    log("Removing Proxy")
    config.proxy = false;
    config.httpsAgent = globalAgent;
    global.GLOBAL_AGENT_FORCE_GLOBAL_AGENT=false;
    if(global.GLOBAL_AGENT) {
        delete global.GLOBAL_AGENT.HTTP_PROXY
    }
}

//Add logging to a http proxy agent
function addLogging(agent: any) {
    if(agent && agent.on) {
        agent.on('keylog', (line : string, tlsSocket: TLSSocket) => {
            try {
                log!("line: %s tlsSocket: %o", line, debugTlsSocket(tlsSocket));
            } catch (err) {
                log!("Unable to trace line: %s with error: %o", line, err)
            }
        });
    }

}

//Add network logging to the config's agent.
export function addLogger(config: AxiosRequestConfig) {
    if(!config.httpsAgent) {
        config.httpsAgent=globalAgent;
    }
    addLogging(config.httpsAgent);
}


export interface AttemptAwareConfig extends AxiosRequestConfig {
    attempts?: number,
    useProxyFirst?: boolean
}

//Adds proxy on/off switch to an axios instance
export async function addInterceptor(instance : AxiosInstance) {
    instance.interceptors.response.use(undefined, async (err) => {
        log("Error encountered.")
        if(err instanceof AxiosError  && hasProxy()) {
            const config = err.config as AttemptAwareConfig;
            if(config.attempts === undefined) {
                config.attempts=0;
            } else {
                config.attempts++;
            }
            log("Failed attempt %s with code for config %s",config.attempts,err.code,err.config)
            if(global.GLOBAL_AGENT_FORCE_GLOBAL_AGENT) {
                removeProxy(config);
            } else {
                addProxy(config);
            }
            if(config.attempts <= 2 && hasProxy() && err instanceof AxiosError && err.code && err.code.search(/[ECONNRESET|ECONNREFUSED|ENOENT]/) >= 0) {
                console.log(`Download failed for file ${config.url}. Attempting ${global.GLOBAL_AGENT_FORCE_GLOBAL_AGENT ? 'without' : 'with'} proxy. Previous Error: ${err}`)
                return await instance(err.config!);
            }
        }
        throw err;
    })
}