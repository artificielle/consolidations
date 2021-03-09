import { Tab, Tabs, TabList, TabPanel } from 'react-tabs'
import { Workspace } from './workspace'

export const Consolidations = () => {
  const topics = ['合并利润表', '合并现金流量表', '合并资产负债表']
  return (
    <div>
      <Tabs>
        <TabList>
          {topics.map((topic, i) => (
            <Tab key={i}>{topic}</Tab>
          ))}
        </TabList>
        {topics.map((topic, i) => (
          <TabPanel key={i}>
            <Workspace topic={topic}></Workspace>
          </TabPanel>
        ))}
      </Tabs>
    </div>
  )
}
