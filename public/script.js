document.getElementById('upload-form').addEventListener('submit', async function (e) {
    e.preventDefault();
    
    const fileInput = document.getElementById('excel-file');
    const numFrases = parseInt(document.getElementById('num-frases').value);
    const resultDiv = document.getElementById('result');

    if (fileInput.files.length === 0) {
        alert("Por favor, selecione um arquivo Excel.");
        return;
    }

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);
    formData.append('numFrases', numFrases);

    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        if (response.ok) {
            const data = await response.json();

            console.log("Dados recebidos:", data);

            const table = document.createElement('table');
            table.innerHTML = `
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Matrícula</th>
                        <th>Regional</th>
                        <th>Frase</th>
                    </tr>
                </thead>
                <tbody>
                    ${data.rows.map(row => `
                        <tr>
                            <td>${row['Nome'] || ''}</td>
                            <td>${row['Matrícula'] || ''}</td>
                            <td>${row['Regional'] || ''}</td>
                            <td>${row['Frase'] || ''}</td>
                        </tr>
                    `).join('')}
                </tbody>
            `;
            resultDiv.innerHTML = '';
            resultDiv.appendChild(table);
        } else {
            const errorText = await response.text();
            resultDiv.textContent = `Erro ao processar o arquivo: ${errorText}`;
        }
    } catch (error) {
        console.error('Erro:', error);
        resultDiv.textContent = "Erro ao processar o arquivo.";
    }
});
