const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const readline = require("readline");

function askQuestion(query) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve) =>
    rl.question(query, (answer) => {
      rl.close();
      resolve(answer);
    })
  );
}

async function processExcelAndSQL() {
  const fichierExcel = await askQuestion(
    "Veuillez entrer le chemin du fichier Excel (.xlsx) : "
  );
  const dossierSortie = await askQuestion(
    "Veuillez entrer le chemin du dossier de sortie pour les fichiers SQL : "
  );
  const codeClient = await askQuestion(
    "Veuillez entrer le code client (account_id) : "
  );

  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(fichierExcel);

    const sheet = workbook.getWorksheet(1);
    const valeurs = [];

    sheet.eachRow((row) => {
      const firstCellValue = row.getCell(1).value;
      valeurs.push(firstCellValue || "");
    });

    console.log("Valeurs extraites :", valeurs);

    if (
      !fs.existsSync(dossierSortie) ||
      !fs.lstatSync(dossierSortie).isDirectory()
    ) {
      console.log("Le chemin spécifié n'est pas un répertoire valide.");
      return;
    }

    const requetesCart = valeurs.map((valeur) => {
      return `INSERT IGNORE INTO cart (customer_code, account_id) VALUES ('${valeur}', ${codeClient});`;
    });

    const requetesCustomerAccount = valeurs.map((valeur) => {
      return `INSERT IGNORE INTO customer_account (customer_code, account_id) VALUES ('${valeur}', ${codeClient});`;
    });

    const fichierCart = path.join(dossierSortie, "cart_insert.sql");
    const fichierCustomerAccount = path.join(
      dossierSortie,
      "customer_account_insert.sql"
    );

    fs.writeFileSync(fichierCart, requetesCart.join("\n"), "utf-8");
    fs.writeFileSync(
      fichierCustomerAccount,
      requetesCustomerAccount.join("\n"),
      "utf-8"
    );

    console.log(
      `Les requêtes pour 'cart' ont été générées dans le fichier : ${fichierCart}`
    );
    console.log(
      `Les requêtes pour 'customer_account' ont été générées dans le fichier : ${fichierCustomerAccount}`
    );
  } catch (error) {
    console.error(
      "Erreur lors de la lecture du fichier Excel ou de l'écriture des fichiers SQL :",
      error
    );
  }
}

processExcelAndSQL().catch((err) => {
  console.error("Erreur lors de l'exécution du script :", err);
});
