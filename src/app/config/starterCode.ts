const base = (host: string) =>
        `await ${host}.run(async context => {
            /**
             * Insert your ${host} code here
             */
            await context.sync();
        });`;

export default (host: string) => {
    switch (host) {
        case 'Outlook':
            // Outlook doesn't use RichAPI and has an empty sample
            return (
        `/**
         * Insert your ${host} code here
         */`
            );
        case 'PowerPoint':
            // PowerPoint doesn't use RichAPI and has it's own sample
            return (
        `/**
         * Insert your ${host} code here
         */
        Office.context.document.setSelectedDataAsync('Hello World!', {
            coercionType: Office.CoercionType.Text
        }, result => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            }
        });`
            );
        default:
            return base(host);
    }
};