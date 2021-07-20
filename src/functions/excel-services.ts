
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

export const writeDataToCell = async (cellAddress: string) => {
    Excel.run(async (context: Excel.RequestContext) => {
        const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
        range.load({values: true});
        await context.sync();
        range.values = [['data1', 'data2'], ['data3', 'data4']];
        await context.sync();
    });
}