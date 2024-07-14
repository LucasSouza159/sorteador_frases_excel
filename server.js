const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Servir arquivos estáticos da pasta 'public'
app.use(express.static(path.join(__dirname, 'public')));

// Rota para upload do arquivo Excel
app.post('/upload', upload.single('file'), (req, res) => {
    const file = req.file;
    const numFrases = parseInt(req.body.numFrases, 10);

    if (!file) {
        console.error("Nenhum arquivo enviado.");
        return res.status(400).send('Nenhum arquivo enviado.');
    }

    try {
        const workbook = xlsx.readFile(file.path);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        let data = xlsx.utils.sheet_to_json(sheet);

        console.log("Data lida do Excel:", data);

        // Mapeamento de colunas para garantir a compatibilidade
        data = data.map(row => ({
            Nome: row['Lucas'],
            Matrícula: row['54015'],
            Regional: row['Sede'],
            Frase: row['Olá mundo']
        }));

        const selectedRows = [];

        for (let i = 0; i < numFrases; i++) {
            const randomIndex = Math.floor(Math.random() * data.length);
            selectedRows.push(data[randomIndex]);
            data.splice(randomIndex, 1);  // Remove a linha sorteada para evitar repetição
        }

        console.log("Selected Rows:", selectedRows);

        res.json({ rows: selectedRows });

        // Remove o arquivo temporário após o processamento
        fs.unlink(file.path, (err) => {
            if (err) {
                console.error("Erro ao remover o arquivo temporário:", err);
            } else {
                console.log("Arquivo temporário removido.");
            }
        });
    } catch (error) {
        console.error("Erro ao processar o arquivo:", error);
        res.status(500).send('Erro ao processar o arquivo.');
    }
});

app.listen(3000, () => {
    console.log('Servidor rodando na porta 3000');
});
