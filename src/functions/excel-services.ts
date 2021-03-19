
export const callExcel = async (invocation: CustomFunctions.StreamingInvocation<string[][]>) => {
    Excel.run( async (context: Excel.RequestContext) => {
        console.log('callExcel')
        context.workbook.onSelectionChanged.add(async (arg): Promise<void> => {
            console.log(arg);
        })
        let range = context.workbook.getSelectedRange();
        range.load({address: true});
        await context.sync();
        console.log('context sync')
        invocation.setResult([[range.address]]);
    })
    
}