import pyarrow as pa
import xlntpyarrow as xpa

print(xpa)

COLUMN_TYPE_FIELD = {
    xpa.Cell.Type.Number: pa.float64,
    xpa.Cell.Type.SharedString: pa.string,
    xpa.Cell.Type.InlineString: pa.string,
    xpa.Cell.Type.FormulaString: pa.string,
    xpa.Cell.Type.Error: pa.string,
    xpa.Cell.Type.Boolean: pa.bool_,
    xpa.Cell.Type.Date: pa.date32,
    xpa.Cell.Type.Empty: pa.string,
}

def xlsx2arrow(io, sheetname):
    reader = xpa.StreamingWorkbookReader()
    reader.open(io)
    print('after open')

    print('before titles')
    sheet_titles = reader.sheet_titles()
    print('after titles', sheet_titles)
    sheet_title = sheet_titles[0]

    if sheetname is not None:
        if isinstance(sheetname, int):
            sheet_title = sheet_titles[sheetname]
        elif isinstance(sheetname, str):
            sheet_title = sheetname

    print('before begin', sheet_title)
    reader.begin_worksheet(sheet_title)
    print('after begin', sheet_title)

    column_names = []
    fields = []
    batches = []

    while reader.has_cell():
        print('read_cell')
        cell = reader.read_cell()
        type = cell.data_type()

        if cell.row() == 1:
            column_names.push_back(cell.value_string())
            continue
        elif cell.row() == 2:
            column_name = column_names[cell.column() - 1]
            fields.append(pa.Field(column_name, COLUMN_TYPE_FIELD[type]()))
            continue
        elif schema is None:
            schema = pa.schema(fields)

        batch = xpa.read_batch(schema, 0)
        print(batch)
        batches.append(batch)

        break

    reader.end_worksheet()

    return pa.Table.from_batches(batches)

if __name__ == '__main__':
    file = open('tmp.xlsx', 'rb')
    print(xlsx2arrow(file, 'Sheet1'))
