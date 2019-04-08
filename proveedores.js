const rp = require('request-promise');
const $ = require('cheerio');
const xl = require('excel4node');
const XLSX = require('xlsx');
const workbook = XLSX.readFile('excel/envio.xlsx');
const sheet_name_list = workbook.SheetNames;

let lstDatosProveedores = [];

function cargarExcel() {
    let lstRut = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

    for (const elemento of lstRut) {
       getDataProveedores(elemento.rut);
    }
}

function getDataProveedores(rut) {

    const url = 'http://webportal.mercadopublico.cl/proveedor/' + rut;
    rp(url)
        .then(function (html) {
            return {
                sitioWeb: $('#lblLinkSitioWeb', html).text(),
                personaContacto: $('#lblPersonaContacto', html).text(),
                telefonoContacto: $('#lblTelefonoContacto', html).text(),
                mail: $('#lblMail', html).text(),
                direccion: $('#lblDireccion', html).text()
              };
        })
        .catch(function (err) {
            //handle error
        }).then(function(dato) {
            lstDatosProveedores.push(dato);
        })
}

function crearExcel(){

    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('Sheet 1');
    let style = wb.createStyle({
        font: {
            color: '#FF0800',
            size: 12,
        }
    });

    for (let index = 0; index < lstDatosProveedores.length; index++) {
        const element = lstDatosProveedores[index];

        ws.cell(index, 1)
        .string(element.sitioWeb)
        .style(style);

        ws.cell(index, 2)
        .string(element.personaContacto)
        .style(style);

        ws.cell(index, 3)
        .string(element.telefonoContacto)
        .style(style);

        ws.cell(index, 4)
        .string(element.mail)
        .style(style);

        ws.cell(index, 5)
        .string(element.direccion)
        .style(style);
    }
    
    wb.write('excel/respuesta.xlsx');
}

cargarExcel();

crearExcel();