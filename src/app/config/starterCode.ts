/**
 * Generates an empty RichAPI-based starter template for the provided host
 * 
 * @param host The host name
 */
const base = (host: string) =>
        `return ${host}.run(context => {
            /**
             * Insert your ${host} code here
             */
            return context.sync();
        });`;

/**
 * Generates a starter code snippet for the provided host.
 * This snippet may be host-specific if the provide host is recognized,
 * otherwise it will be a generic base snippet.
 * 
 * @param host The host name
 */
export default (host: string) => {
    switch (host) {
        case 'Excel':
            return (
        `return ${host}.run(context => {
            /**
             * Insert your ${host} code here
             */
            const range = context.workbook.getSelectedRange();
            
            // Read the range address
            range.load('address');

            // Update the fill color
            range.format.fill.color = 'yellow';

            return context.sync().then(() => 
                console.log(\`The range address was \${range.address}.\`);
            );
        });`
            );
        case 'Word':
            return (
        `return ${host}.run(context => {
            /**
             * Insert your ${host} code here
             */
            const range = context.document.getSelection();
            
            // Read the range text
            range.load('text');

            // Update font color
            range.font.color = 'red';

            return context.sync().then(() => 
                console.log(\`The selected text was \${range.text}.\`);
            );
        });`
            );
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