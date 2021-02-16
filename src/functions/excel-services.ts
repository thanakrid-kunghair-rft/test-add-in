export const callExcel = async (invocation: CustomFunctions.StreamingInvocation<string[][]>) => {
    let context = new Excel.RequestContext();
    let range = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
    range.load();
    await context.sync();
    invocation.setResult(range.values);
}