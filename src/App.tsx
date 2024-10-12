import reactLogo from './assets/react.svg'
import viteLogo from '/vite.svg'
import './App.css'
import {Button, Space, Table, Upload} from "antd";
import {exportExcel, uploadExcel} from "./ExcelUtils.ts";
import {useState} from "react";


export type DataType = {
    key: string;
    app: string;
    name: string;
    works: string;
}

const dataSource = [
    {
        key: '1',
        app: "掘金",
        name: '程序饲养员',
        works: '22篇',
    },
    {
        key: '1',
        app: "知乎",
        name: '程序饲养员',
        works: '23篇',
    },
];

const columns = [
    {
        title: 'APP',
        dataIndex: 'app',
        key: 'app',
    },
    {
        title: '名称',
        dataIndex: 'name',
        key: 'name',
    },
    {
        title: '作品数量',
        dataIndex: 'works',
        key: 'works',
    },
];

function App() {

    const [tableData, setTableData] = useState(dataSource)

    return (
        <>
            <div>
                <a href="https://vitejs.dev" target="_blank">
                    <img src={viteLogo} className="logo" alt="Vite logo"/>
                </a>
                <a href="https://react.dev" target="_blank">
                    <img src={reactLogo} className="logo react" alt="React logo"/>
                </a>
            </div>
            <h1>Vite + React + ExcelJS</h1>
            <Table dataSource={tableData} columns={columns} pagination={false}/>
            <div className="card">
                <Space size={'large'}>
                    <Button onClick={() => exportExcel(tableData)}>下载Excel</Button>
                    <Upload
                        multiple
                        showUploadList={false}
                        action="/"
                        beforeUpload={async (file) => {
                            const excelData = await uploadExcel(file);
                            setTableData([...tableData, ...excelData])
                            return false;
                        }}
                    >
                        <Button>上传Excel</Button>
                    </Upload>
                </Space>
                <p>
                    Edit <code>src/App.tsx</code> and save to test HMR
                </p>
            </div>
            <p className="read-the-docs">
                Click on the Vite and React logos to learn more
            </p>
        </>
    )
}

export default App
