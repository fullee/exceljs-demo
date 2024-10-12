import {DataType} from "./App.tsx";
import * as ExcelJS from 'exceljs'

export const uploadExcel = async (file: File) => {
    const arrayBuffer =  await file.arrayBuffer()

    {
        const tableData: DataType[] = [];
        const workbook = new ExcelJS.Workbook();
        try {
            await workbook.xlsx.load(arrayBuffer);
            // 获取第一个工作表
            const worksheet = workbook.getWorksheet(1);

            // 读取工作表中的数据
            worksheet?.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                console.log(`Row ${rowNumber}:`, row.values);
                // 去掉表头
                if (rowNumber > 1) {
                    tableData.push({
                        key: rowNumber.toString(),
                        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
                        // @ts-expect-error
                        app: row.values[1].trim(),
                        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
                        // @ts-expect-error
                        name: row.values[2].trim(),
                        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
                        // @ts-expect-error
                        works: row.values[3].trim()
                    })
                }
            });
            // console.log('employeeNos', worksheet.rowCount, employeeNos.length);
            // if ((worksheet.rowCount - 1) !== employeeNos.length) {
            //   message.error('导入失败，请检查Excel文件数据是否正确');
            //   return [];
            // }
        } catch (error) {
            console.error('Error loading workbook:', error);
        }

        console.log(tableData);
        return tableData;
    }
}

export const exportExcel = (data: DataType[] ) => {
    const headerStyle = {
        font: {
            name: 'Arial',
            family: 4,
            size: 12,
            bold: true,
            // color: { argb: 'FF0000' }
        },
        fill: {
            type: 'pattern',
            pattern: 'solid',
            // fgColor: { argb: 'FFFF00' },
            // bgColor: { argb: 'FFFF00' }
        },
        alignment: {
            vertical: 'middle',
            horizontal: 'center'
        },
        border: {
            top: {style: 'thin', color: {argb: '000000'}},
            left: {style: 'thin', color: {argb: '000000'}},
            bottom: {style: 'thin', color: {argb: '000000'}},
            right: {style: 'thin', color: {argb: '000000'}}
        }
    };

    const headerTitle = ['APP', '名称', '作品数']

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("Sheet1")
    ws.addRow(headerTitle)
        .eachCell((cell) => {
            // eslint-disable-next-line @typescript-eslint/ban-ts-comment
            // @ts-expect-error
            cell.style = headerStyle
        })
    data.forEach(it=> {
        ws.addRow(Object.values({
            app: it.app,
            name: it.name,
            works: it.works
        }))

    })
    // .eachCell((cell, index)=> {
    //   if (index === 1) {
    //     cell.numFmt = '0'
    //   }
    // })

    ws.columns = headerTitle.map((header) => ({
        header, key: header, width: 20
    }))

    workbook.xlsx.writeBuffer()
        .then(buffer => {
            // 创建 Blob 对象
            const blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

            // 创建下载链接
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `exceljs.xlsx`;
            a.click();

            // 清理 URL
            setTimeout(() => {
                window.URL.revokeObjectURL(url);
                a.remove();
            }, 100);
        })
        .catch(err => console.error('Error creating file:', err));
}
