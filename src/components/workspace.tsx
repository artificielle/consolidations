import { useCallback, useEffect, useState, Dispatch, SetStateAction } from 'react'
import { Tab, Tabs, TabList, TabPanel } from 'react-tabs'
import { useDropzone } from 'react-dropzone'
import { Worksheet } from './worksheet'
import { saveAs } from 'file-saver'
import excel from 'exceljs'

import 'react-tabs/style/react-tabs.css'
import './workspace.css'

type WorkspaceProps = {
  topic: string
}

export const Workspace = ({ topic }: WorkspaceProps) => {
  const [template, setTemplate] = useState<excel.Workbook>()
  const [consolidations, setConsolidations] = useState<excel.Worksheet[]>([]) // 合并, 合并分公司
  const [subsidiaries, setSubsidiaries] = useState<excel.Worksheet[]>([]) // 子公司
  const [branches, setBranches] = useState<excel.Worksheet[]>([]) // 分公司

  useEffect(() => {
    const loadTemplate = async () => {
      try {
        const response = await fetch('/xlsx-templates/20XX' + topic + '（模板）.xlsx')
        const buffer = await response.arrayBuffer()
        const template = new excel.Workbook()
        await template.xlsx.load(buffer)
        const consolidations = [template.worksheets[0], template.worksheets[6]]
        if (topic === '合并现金流量表') consolidations[1] = template.worksheets[7]
        setTemplate(template)
        setConsolidations(consolidations)
        console.log('load template completed', template)
      } catch (error) {
        console.error(error)
        alert('模板文件错误')
      }
    }
    loadTemplate()
  }, [topic])

  const loadWith = (setState: Dispatch<SetStateAction<excel.Worksheet[]>>) => async (files: File[]) => {
    try {
      const worksheets = await Promise.all(
        files.map(async (file) => {
          const buffer = await readFile(file)
          const workbook = new excel.Workbook()
          await workbook.xlsx.load(buffer)
          return workbook.worksheets[0]
        }),
      )
      setState((w) => [...w, ...worksheets])
    } catch (error) {
      console.error('read xlsx file', error)
      alert('文件格式错误')
    }
  }

  const loadSubsidiariesDropzone = useDropzone({ onDrop: loadWith(setSubsidiaries) })
  const loadBranchesDropzone = useDropzone({ onDrop: loadWith(setBranches) })

  const save = async () => {
    try {
      if (!template) throw new Error()
      const buffer = await template.xlsx.writeBuffer()
      saveAs(new Blob([buffer]), topic + '.xlsx')
    } catch (error) {
      console.error('write xlsx file', error)
      alert('模板文件错误')
    }
  }

  const consolidate = useCallback(() => {
    // TODO
    const sum = (numbers: number[]) => numbers.reduce((x, y) => x + y, 0)
    const validateCell = (cell: excel.Cell) => {
      if (cell.type !== excel.ValueType.Number) throw new Error(`${cell.worksheet.name} ${cell.address}`)
      return cell
    }
    // const consolidateCells = (worksheet: excel.Worksheet, target: string, sources: string[], sourcesNeg: string[]) => {
    //   worksheet.getCell(target).value =
    //     sum(expandAddresses(sources).map((address) => validateCell(worksheet.getCell(address)).value as number)) -
    //     sum(expandAddresses(sourcesNeg).map((address) => validateCell(worksheet.getCell(address)).value as number))
    // }
    const consolidateWorksheets = (target: excel.Worksheet, sources: excel.Worksheet[], addresses: string[]) => {
      for (const address of expandAddresses(addresses)) {
        target.getCell(address).value = sum(
          sources.map((worksheet) => validateCell(worksheet.getCell(address)).value as number),
        )
      }
    }
    const consolidateBranches = (addresses: string[]) => {
      if (!branches.length) return
      consolidateWorksheets(consolidations[1], branches, addresses)
    }
    const consolidateSubsidiaries = (addresses: string[]) => {
      if (!subsidiaries.length) return
      consolidateWorksheets(consolidations[0], [consolidations[1], ...subsidiaries], addresses)
    }
    // const consolidateBranchesInside = (target: string, sources: string[], sourcesNeg: string[] = []) => {
    //   consolidateCells(consolidations[1], target, sources, sourcesNeg)
    // }
    // const consolidateSubsidiariesInside = (target: string, sources: string[], sourcesNeg: string[] = []) => {
    //   consolidateCells(consolidations[0], target, sources, sourcesNeg)
    // }
    try {
      if (topic === '合并利润表') {
        consolidateBranches(['B5:C21', 'B23:C24', 'B26:C26', 'B28:C33'])
        consolidateSubsidiaries(['B5:C21', 'B23:C24', 'B26:C26'])
      }
      if (topic === '合并现金流量表') {
        consolidateBranches(['B6:C8', 'B10:C13', 'B17:C21', 'B23:C26', 'B30:C32', 'B34:C36', 'B41'])
        consolidateSubsidiaries(['B6:C8', 'B10:C13', 'B17:C21', 'B23:C26', 'B30:C32', 'B34:C36', 'B41'])
      }
      if (topic === '合并资产负债表') {
        consolidateBranches(['B6:B10', 'B11:B20', 'B23:B30', 'B31:B40'])
        consolidateBranches(['E6:E10', 'E11:E20', 'E23:E30', 'E31:E40', 'E41:E47'])
        // consolidateBranchesInside('B21', ['B6:B13', 'B16:B20'])
        // consolidateBranchesInside('B41', ['B23:B40'])
        // consolidateBranchesInside('B47', ['B21', 'B41'])
        // consolidateBranchesInside('E21', ['E6:E15', 'E18:E20'])
        // consolidateBranchesInside('E33', ['E23:E24', 'E27:E32'])
        // consolidateBranchesInside('E34', ['E21', 'E33'])
        // consolidateBranchesInside('E46', ['E36:E37', 'E42:E45', 'E40'], ['E41'])
        // consolidateBranchesInside('E47', ['E34', 'E46'])
        // consolidateBranchesInside('H43', ['E45'], ['F45'])
        consolidateSubsidiaries(['B6:B10', 'B11:B20', 'B23:B30', 'B31:B40'])
        consolidateSubsidiaries(['E6:E10', 'E11:E20', 'E23:E30', 'E31:E40', 'E41:E47'])
        // consolidateSubsidiariesInside('B21', ['B6:B13', 'B16:B20'])
        // consolidateSubsidiariesInside('B41', ['B23:B40'])
        // consolidateSubsidiariesInside('B47', ['B21', 'B41'])
        // consolidateSubsidiariesInside('E21', ['E6:E15', 'E18:E20'])
        // consolidateSubsidiariesInside('E33', ['E23:E24', 'E27:E32'])
        // consolidateSubsidiariesInside('E34', ['E21', 'E33'])
        // consolidateSubsidiariesInside('E46', ['E36:E37', 'E42:E45', 'E40'], ['E41'])
        // consolidateSubsidiariesInside('E47', ['E34', 'E46'])
        // consolidateSubsidiariesInside('H43', ['E45'], ['F45'])
      }
    } catch (error) {
      console.error('consolidate', error)
      alert('文件格式错误: ' + error)
    }
    setConsolidations([...consolidations])
    console.log('consolidate completed', consolidations, subsidiaries, branches)
  }, [topic, consolidations, subsidiaries, branches])

  return (
    <div>
      <div className="workspace-dropzone-container">
        <div {...loadSubsidiariesDropzone.getRootProps()}>
          <input {...loadSubsidiariesDropzone.getInputProps()} />
          <span>添加分公司</span>
        </div>
        <div {...loadBranchesDropzone.getRootProps()}>
          <input {...loadBranchesDropzone.getInputProps()} />
          <span>添加子公司</span>
        </div>
        <div onClick={save}>
          <span>下载合并表</span>
        </div>
      </div>
      <Tabs onSelect={consolidate}>
        <TabList>
          {consolidations.map((_, i) => (
            <Tab key={i}>{[topic, topic.replace('合并', '合并分公司')][i]}</Tab>
          ))}
          {subsidiaries.map((_, i) => (
            <Tab key={i}>子公司 {i + 1}</Tab>
          ))}
          {branches.map((_, i) => (
            <Tab key={i}>分公司 {i + 1}</Tab>
          ))}
        </TabList>
        {consolidations.map((worksheet, i) => (
          <TabPanel key={i}>
            <Worksheet worksheet={worksheet}></Worksheet>
          </TabPanel>
        ))}
        {subsidiaries.map((worksheet, i) => (
          <TabPanel key={i}>
            <Worksheet worksheet={worksheet}></Worksheet>
          </TabPanel>
        ))}
        {branches.map((worksheet, i) => (
          <TabPanel key={i}>
            <Worksheet worksheet={worksheet}></Worksheet>
          </TabPanel>
        ))}
      </Tabs>
    </div>
  )
}

const readFile = async (file: File) =>
  new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (event) => {
      try {
        resolve(event.target!.result! as ArrayBuffer)
      } catch (error) {
        reject(error)
      }
    }
    reader.readAsArrayBuffer(file)
  })

const expandAddresses = (ss: string[]) => ss.flatMap(expandAddress)

const expandAddress = (s: string): string[] => {
  if (!s.includes(':')) return [s]
  const m = s.match(/^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$/)
  if (!m || m.length < 5) throw new Error('Invalid range: ' + s)
  const t = Number(m[2])
  const b = Number(m[4])
  const l = m[1].charCodeAt(0)
  const r = m[3].charCodeAt(0)
  const indexes = (a: number, z: number) => Array.from({ length: z - a + 1 }).map((_, i) => a + i)
  return indexes(l, r).flatMap((c) => indexes(t, b).map((r) => String.fromCharCode(c) + String(r)))
}
