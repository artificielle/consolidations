import excel from 'exceljs'
import './worksheet.css'

type WorksheetProps = { worksheet: excel.Worksheet }

export const Worksheet = ({ worksheet }: WorksheetProps) => {
  const { table, size } = makeTable(worksheet)
  return (
    <div className="worksheet-container">
      <table>
        <tbody>
          {makeTableIndexes(size[0]).map((r) => (
            <tr key={r}>
              {makeTableIndexes(size[1]).map((c) => (
                <td key={c}>
                  <div>{renderCell(table?.[r]?.[c])}</div>
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

const renderCell = (cell: excel.Cell) => {
  if (!cell) return undefined
  if (cell.master !== cell) return undefined
  if (cell.type === excel.ValueType.Formula) return renderFormulaCell(cell)
  return cell.text
}

const renderFormulaCell = (cell: excel.Cell) => {
  const result = (cell?.value as excel.CellFormulaValue)?.result as excel.CellErrorValue | undefined
  return result?.error ?? result
}

const makeTableIndexes = (length: number) => Array.from({ length }).map((_, i) => i + 1)

const makeTable = (worksheet: excel.Worksheet) => {
  const table: excel.Cell[][] = []
  const size: [number, number] = [0, 0]
  worksheet.eachRow((row, r) => {
    size[0] = Math.max(size[0], r)
    row.eachCell((cell, c) => {
      size[1] = Math.max(size[1], c)
      table[r] ??= []
      table[r][c] = cell
    })
  })
  return { table, size }
}
