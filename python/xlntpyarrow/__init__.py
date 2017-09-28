import pyarrow as pa
import xlntpyarrow.lib as xpa

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

def cell_to_pyarrow_array(cell, type):
    if cell.data_type() == xpa.Cell.Type.Number:
        return pa.array([cell.value_double()], type)
    elif cell.data_type() == xpa.Cell.Type.SharedString:
        return pa.array([cell.value_string()], type)
    elif cell.data_type() == xpa.Cell.Type.InlineString:
        return pa.array([cell.value_string()], type)
    elif cell.data_type() == xpa.Cell.Type.FormulaString:
        return pa.array([cell.value_string()], type)
    elif cell.data_type() == xpa.Cell.Type.Error:
        return pa.array([cell.value_string()], type)
    elif cell.data_type() == xpa.Cell.Type.Boolean:
        return pa.array([cell.value_bool()], type)
    elif cell.data_type() == xpa.Cell.Type.Date:
        return pa.array([cell.value_unsigned_int()], type)
    elif cell.data_type() == xpa.Cell.Type.Empty:
        return pa.array([cell.value_string()], type)

def xlsx2arrow(io, sheetname):
    reader = xpa.StreamingWorkbookReader()
    reader.open(io)

    sheet_titles = reader.sheet_titles()
    sheet_title = sheet_titles[0]

    if sheetname is not None:
        if isinstance(sheetname, int):
            sheet_title = sheet_titles[sheetname]
        elif isinstance(sheetname, str):
            sheet_title = sheetname

    reader.begin_worksheet(sheet_title)

    column_names = []
    fields = []
    batches = []
    schema = None
    first_batch = []
    max_column = 0

    while reader.has_cell():
        if schema is None:
            cell = reader.read_cell()
            type = cell.data_type()

            if cell.row() == 1:
                column_names.append(cell.value_string())
                max_column = max(max_column, cell.column())
                continue
            elif cell.row() == 2:
                column_name = column_names[cell.column() - 1]
                if type == xpa.Cell.Type.Number and cell.format_is_date():
                    fields.append(pa.field(column_name, pa.date32))
                else:
                    fields.append(pa.field(column_name, COLUMN_TYPE_FIELD[type]()))
                first_batch.append(cell_to_pyarrow_array(cell, fields[-1].type))
                if cell.column() == max_column:
                    schema = pa.schema(fields)
                    print(schema)
                    batches.append(pa.RecordBatch.from_arrays(first_batch, column_names))
                continue

        batches.append(reader.read_batch(schema, 10000))

    reader.end_worksheet()

    return pa.Table.from_batches(batches)

if __name__ == '__main__':
    file = open('tmp.xlsx', 'rb')
    table = xlsx2arrow(file, 'Sheet1')
    print(table.to_pandas())
