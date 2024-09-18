const fs = require('fs');
const ExcelJS = require('exceljs');
const readline = require('readline');

async function processExcelAndSQL() {
    const fichierExcel = 'data.xlsx';
    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.readFile(fichierExcel);

    const sheet = workbook.getWorksheet(1); 

    const valeurs = [];
    sheet.eachRow((row, rowNumber) => {
        const firstCellValue = row.getCell(1).value; 
        valeurs.push(firstCellValue || ''); x
    });

    console.log('Valeurs extraites :', valeurs);

   
    const fichierSQL = 'insert.sql';
    let requetesSQL = fs.readFileSync(fichierSQL, 'utf-8').split('\n');

    function remplacerTmpDansRequetes(requetes, valeurs) {
        if (valeurs.length < requetes.length) {
            console.log("Attention : Moins de valeurs que de requêtes. Certaines requêtes ne seront pas modifiées.");
        }

        return requetes.map((requete, index) => {
            if (index < valeurs.length) {
                return requete.replace('tmp', valeurs[index] || ''); 
            }
            return requete; 
        });
    }

    const nouvellesRequetes = remplacerTmpDansRequetes(requetesSQL, valeurs);

    const contenuFichierFinal = nouvellesRequetes.join('\n');

    const fichierSortie = 'output.sql';
    fs.writeFileSync(fichierSortie, contenuFichierFinal, 'utf-8');

    console.log(`Les requêtes ont été mises à jour et enregistrées dans le fichier ${fichierSortie}.`);
}

processExcelAndSQL().catch(err => {
    console.error('Erreur lors de l\'exécution du script :', err);
});
