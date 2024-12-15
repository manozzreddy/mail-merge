import { TemplateHandler } from "easy-template-x";
import readXlsxFile from 'read-excel-file'
import JSZip from "jszip";
import { saveAs } from "file-saver";

const createDocxBtn = document.getElementById("create-docx-button");

createDocxBtn?.addEventListener("click", async () => {
    const template = await getDocxTemplate();
    const data = await getData() as Array<Object>;

    const handler = new TemplateHandler();
    const zip = new JSZip();

    for (const row of data) {
        const doc = await handler.process(template, row);

        // Add the generated document to the zip file
        zip.file(`${row['First_Name']}.docx`, doc);        
    }

    // Generate the zip file and save it
    zip.generateAsync({ type: 'blob' }).then((content) => {
        saveAs(content, 'documents.zip');
    });
});


function getDocxTemplate() {
    return new Promise((resolve, reject) => {
        const fileInput = document.getElementById('upload-template-button');
        const file = fileInput.files[0];

        if (file) {
            const reader = new FileReader();

            reader.onload = function (event) {
                const blob = new Blob([event.target.result]);
                resolve(blob);
            };

            reader.onerror = function (event) {
                reject(new Error('Error reading file'));
            };

            reader.readAsArrayBuffer(file);
        } else {
            reject(new Error('No file selected'));
        }
    });
}


function getData() {
    return new Promise((resolve, reject) => {
        const fileInput = document.getElementById('upload-excel-button');
        const file = fileInput.files[0];

        if (file) {
            return readXlsxFile(file).then((rows) => {
                // Check if there is at least one row for header
                if (rows.length === 0) {
                    resolve(data);
                } else {

                    // Extract header row
                    const headers = rows[0];

                    // Map the remaining rows to objects using the headers as keys
                    const jsonData = rows.slice(1).map(row => {
                        const obj = {};
                        headers.forEach((header, index) => {
                            obj[header] = row[index];
                        });
                        return obj;
                    });

                    resolve(jsonData);
                };

            }).catch((error: any) => {
                reject(new Error(error));
            });

        } else {
            reject(new Error('No file selected'));
        }
    });
}
