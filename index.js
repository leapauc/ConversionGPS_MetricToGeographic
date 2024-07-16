const xlsx = require('xlsx');
const proj4 = require('proj4');
const express = require('express');
const app = express();
const port = 3000;

// Define the projection
// EPSG20350 => AGD84 / AMG zone 50 
proj4.defs("EPSG:20350", '+proj=tmerc +lat_0=0 +lon_0=117 +k=0.9996 +x_0=500000 +y_0=10000000 +ellps=krass +units=m +no_defs +type=crs');
proj4.defs("EPSG:4326", '+proj=longlat +datum=WGS84 +no_defs'); // WGS84

const epsg20350 = proj4.defs('EPSG:20350');
const wgs84 = proj4.defs('EPSG:4326');

// Fonction pour lire un fichier Excel et retourner les données
function readExcelFile(filePath) {
  try {
    const wb = xlsx.readFile(filePath);
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(ws);
    return data;
  } catch (error) {
    console.error(`Error reading the Excel file: ${error.message}`);
    return null;
  }
}

// Fonction pour convertir les coordonnées
function convertCoordinates(data) {
    return data.map(row => {
      const x = parseFloat(row.X);
      const y = parseFloat(row.Y);
      
      if (!isNaN(x) && !isNaN(y)) {
        const [longitude, latitude] = proj4(epsg20350, wgs84, [x, y]);
        return { ...row, longitude, latitude };
      } else {
        console.warn(`Invalid coordinates for row: ${JSON.stringify(row)}`);
        return { ...row, longitude: null, latitude: null };
      }
    });
  }

// Appeler la fonction pour lire le fichier Excel et convertir les coordonnées
function processExcelFile(filePath) {
  const excelData = readExcelFile(filePath);
  if (excelData) {
    const convertedData = convertCoordinates(excelData);
    return convertedData;
  } else {
    return null;
  }
}

// Exemple d'utilisation dans un endpoint Express
app.get('/convert', (req, res) => {
  const filePath = "./collar.xlsx";
  const result = processExcelFile(filePath);
  if (result) {
    const filteredData = result.map(row => ({
        Hole_ID: row.Hole_ID,
        longitude: row.longitude,
        latitude: row.latitude,
        RL_AHD: row.RL_AHD,
        Depth: row.Depth
      }));
    const newWB = xlsx.utils.book_new();
    // Créer une nouvelle feuille de calcul avec les données résultantes
    const newWS = xlsx.utils.json_to_sheet(filteredData);
    // Ajouter la nouvelle feuille de calcul au workbook
    xlsx.utils.book_append_sheet(newWB, newWS, "Converted Data");
    // Écrire le workbook dans un nouveau fichier Excel
    xlsx.writeFile(newWB, "collar_WGS84.xlsx");

    res.json(result);
    console.table(result);
  } else {
    res.status(500).send('Failed to read or process the Excel file.');
  }
});

app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
  console.log(`GREAT !`);
});
