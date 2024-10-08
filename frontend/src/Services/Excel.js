import { utils, writeFile } from 'xlsx';

export const exportToExcel = (data, fileName) => {
    
    const wb = utils.book_new();
    const ws = utils.json_to_sheet([]);

    
    utils.sheet_add_aoa(ws, [["Uni Bodim Finder - Feedback Report"]], { origin: "F1" });
    
    //MERGE
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { c: 5, r: 0 }, e: { c: 7, r: 0 } });
    ws['!merges'].push({ s: { c: 5, r: 1 }, e: { c: 7, r: 1 } });


    const currentDate = new Date().toLocaleDateString(); 
    ws['F2'] = {
        v: `Report Date: ${currentDate}`, 
        t: 's', 
        s: {
            font: {
                bold: true, 
                sz: 14, 
                color: { rgb: "0000FF" } 
            },
            alignment: {
                vertical: 'center',
                horizontal: 'left',
            }
        }
    };

   
    if (!ws['!cols']) ws['!cols'] = [];

        //  column width
    ws['!cols'][3] = { wch: 25 }; 
    ws['!cols'][1] = { wch: 25 };
    ws['!cols'][2] = { wch: 25 };
    ws['!cols'][7] = { wch: 12 };
    ws['!cols'][8] = { wch: 25 };
    ws['!cols'][9] = { wch: 25 };
    ws['!cols'][10] = { wch: 25 };
    ws['!cols'][11] = { wch: 0 };
    if (!ws['!rows']) ws['!rows'] = [];

        // row height
    ws['!rows'][0] = { hpt: 50 }; 
    ws['!rows'][1] = { hpt: 25 }; 
    ws['!rows'][4] = { hpt: 20 }; 

  
    utils.sheet_add_json(ws, data, { origin: "A5" });
    utils.book_append_sheet(wb, ws, "Feedback");
    writeFile(wb, fileName); 
};
