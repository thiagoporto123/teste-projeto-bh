const fs = require('fs');
const express = require('express');
const session = require('express-session'); // Importar express-session
const { google } = require('googleapis');
const bodyParser = require('body-parser');
const path = require('path');

const app = express();
const port = 3000;

// Configuração do express-session
app.use(session({
    secret: 'seu_segredo_aqui', // Troque por um segredo real
    resave: false,
    saveUninitialized: true,
}));

// Serve arquivos estáticos (HTML, CSS)
app.use(express.static(path.join(__dirname, 'public')));
app.use(bodyParser.urlencoded({ extended: true }));

// Carrega as credenciais da API do Google
const credentials = JSON.parse(fs.readFileSync('credentials.json'));
const { client_secret, client_id, redirect_uris } = credentials.web;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

// Função para conectar ao Google Sheets
function getAuthenticatedSheetsClient() {
    return new Promise((resolve, reject) => {
        fs.readFile('token.json', (err, token) => {
            if (err) return getNewToken(oAuth2Client, resolve);
            oAuth2Client.setCredentials(JSON.parse(token));
            resolve(google.sheets({ version: 'v4', auth: oAuth2Client }));
        });
    });
}

// Rota para a página de login
app.get('/login', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Rota para processar o login
app.post('/login', (req, res) => {
    const { email } = req.body;
    // Armazena o e-mail na sessão
    req.session.email = email; 
    // Redireciona para a página principal após login
    res.redirect('/');
});

// Rota principal
app.get('/', (req, res) => {
    // Verifica se o usuário está logado
    if (!req.session.email) {
        return res.redirect('/login'); // Redireciona para o login se não estiver logado
    }
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Rota para buscar o colaborador
app.get('/buscar-colaborador', async (req, res) => {
    const { matricula } = req.query;
    const sheets = await getAuthenticatedSheetsClient();
    const spreadsheetId = '18IwIenWl-d8ckK8kD_Ck8lQTsZIFWQzwR4LRVVyTSCE'; // ID da planilha para buscar colaborador

    const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: 'E:F', // Colunas a serem lidas
    });

    const rows = response.data.values;
    let colaboradorNome = '';

    if (rows.length) {
        for (const row of rows) {
            if (row[0] === matricula) {
                colaboradorNome = row[1]; // Nome do colaborador na coluna F
                break;
            }
        }
    }

    res.json({ nome: colaboradorNome });
});

// Rota para processar o formulário
app.post('/solicitar-bh', async (req, res) => {
    const { matricula, dataBH, horaInicio, horaFim } = req.body;
    const email = req.session.email; // Pega o e-mail da sessão

    const sheets = await getAuthenticatedSheetsClient();
    const spreadsheetId = '1Iip-bR9Y18qy7zHdTjF7A5jmtmrwnag_46fzjwXoGTU'; // ID da sua planilha

    // Verificar se a matrícula existe
    const verificaResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: '18IwIenWl-d8ckK8kD_Ck8lQTsZIFWQzwR4LRVVyTSCE',
        range: 'E:F', // Colunas a serem lidas
    });

    const verificaRows = verificaResponse.data.values;
    let colaboradorNome = '';
    let matriculaValida = false;

    if (verificaRows.length) {
        for (const row of verificaRows) {
            if (row[0] === matricula) { // Verifica se a matrícula está na coluna E
                matriculaValida = true;
                colaboradorNome = row[1]; // Nome do colaborador na coluna F
                break;
            }
        }
    }

    if (!matriculaValida) {
        return res.send(`<script>alert('Matrícula inválida'); window.history.back();</script>`); // Pop-up para matrícula inválida
    }

    // Se a matrícula for válida, prossegue com o envio
    await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: 'A:H', // Colunas a serem preenchidas
        valueInputOption: 'USER_ENTERED',
        resource: {
            values: [[matricula, colaboradorNome, dataBH, horaInicio, horaFim, new Date().toLocaleDateString(), new Date().toLocaleTimeString(), email]],
        },
    }, (err, result) => {
        if (err) {
            console.log(err);
            res.send('Erro ao enviar dados!');
        } else {
            res.send('Solicitação enviada com sucesso!');
        }
    });
});

// Inicia o servidor
app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
});
