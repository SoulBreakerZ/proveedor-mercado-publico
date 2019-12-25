const $ = require('cheerio');
const xl = require('excel4node');

const XLSX = require('xlsx');
const workbook = XLSX.readFile('excel/envio.xlsx');
const sheet_name_list = workbook.SheetNames;

const http = require('http');

const urls = [];
const lstDatosProveedores = [];

class Proveedor {
    constructor(sitioWeb,personaContacto,telefonoContacto,mail,direccion) {
        this.sitioWeb = sitioWeb;
        this.personaContacto = personaContacto;
        this.telefonoContacto = telefonoContacto;
        this.mail = mail;
        this.direccion = direccion;
    }
}

function cargarExcel() {
    let lstRut = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    for (const element of lstRut) {
        urls.push('http://webportal.mercadopublico.cl/proveedor/' + element.rut);
    }
}

function crearDataProveedores() {
    var completed_requests = 0;

    urls.forEach(function(url) {
        http.get(url, (resp) => {
            let data = '';

            resp.on('data', (chunk) => {
                data += chunk;
            });

            resp.on('end', () => {
                lstDatosProveedores.push(new Proveedor($('#lblLinkSitioWeb', data).text(),
                $('#lblPersonaContacto', data).text(),$('#lblTelefonoContacto', data).text(),$('#lblMail', data).text(),$('#lblDireccion', data).text()));
                if (completed_requests++ == urls.length - 1) {
                    crearExcel();
                }     
            });
    
            }).on("error", (err) => {
                console.log("Error: " + err.message);
            });
    });
}

function crearExcel(){
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('Sheet 1');

    crearCabezeraExcel(wb,ws);

    let style = wb.createStyle({
        font: {
            color: '#000000',
            size: 12,
        }
    });

    let indexExcel = 2;
    for (const value of lstDatosProveedores) {
        
        ws.cell(indexExcel, 1)
        .string(value.sitioWeb)
        .style(style);

        ws.cell(indexExcel, 2)
        .string(value.personaContacto)
        .style(style);

        ws.cell(indexExcel, 3)
        .string(value.telefonoContacto)
        .style(style);

        ws.cell(indexExcel, 4)
        .string(value.mail)
        .style(style);

        ws.cell(indexExcel, 5)
        .string(value.direccion)
        .style(style);

        indexExcel++;
    }
    
    wb.write('excel/respuesta.xlsx');
}

function crearCabezeraExcel(wb,ws){

    let style = wb.createStyle({
        font: {
            color: '#000000',
            size: 12,
        }
    });

    ws.cell(1, 1)
    .string('sitioWeb')
    .style(style);

    ws.cell(1, 2)
    .string('personaContacto')
    .style(style);

    ws.cell(1, 3)
    .string('telefonoContacto')
    .style(style);

    ws.cell(1, 4)
    .string('mail')
    .style(style);

    ws.cell(1, 5)
    .string('direccion')
    .style(style);

}


cargarExcel();
crearDataProveedores();
