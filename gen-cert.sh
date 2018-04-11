if [[ $(uname) == *"MINGW64"* ]]
then
    # This script is intended to run in that Git Bash environment. Note the form for -subj
    echo "Generating RSA key for the root CA and store it in ca.key:"
    openssl genrsa -out ca.key 2048

    echo ""
    echo "Create the self-signed root CA certificate in ca.crt:"

    openssl req -new -x509 -nodes -days 1826 -key ca.key -sha256 -out ca.crt -subj "//C=US\ST=WA\L=Redmond\O=Office\OU=OfficeExtensibility\CN=localhost-ca"

    echo ""
    echo "Create private key for subordinate CA:"
    openssl genrsa -out server.key 4096

    echo ""
    echo "Request a certificate for the subordinate CA:"

    openssl req -new -key server.key -config san.conf -out server.csr -subj "//C=US\ST=WA\L=Redmond\O=Office\OU=OfficeExtensibility\CN=localhost"

    echo ""
    echo "Process the subordinate CA cert request and sign it with the root CA:"
    openssl x509 -req -days 36500 -in server.csr -CA ca.crt -CAkey ca.key -set_serial 01 -out server.crt

    echo ""
    echo "NEXT STEP (required): install the root CA (ca.crt) in your Trusted Root Certification Authorities store."

else
    echo "create certs not with Git Bash env, you'll need to set execute perms: chmod +x gen-cert.sh"
fi