const express = require('express');
const app = express();
const port = 3000;
const fs = require('fs');
const Docxtemplater = require('docxtemplater');
const { promisify } = require('util');

// Função para carregar o modelo do certificado
const loadTemplate = async () => {
    try {
        // Carrega o conteúdo do arquivo do modelo de certificado
        const content = fs.readFileSync('CertificadoRTO.docx', 'binary');
        
        // Cria uma instância do Docxtemplater e carrega o conteúdo do modelo
        const doc = new Docxtemplater();
        doc.loadZip(content);
        
        // Retorna o objeto docxtemplater carregado
        return doc;
    } catch (error) {
        // Trata qualquer erro que ocorra durante a leitura do arquivo
        console.error('Erro ao carregar o modelo de certificado:', error);
        throw error; // Rejeita a promessa com o erro
    }
};

// Função para preencher o modelo do certificado com os dados da empresa e do usuário
const generateCertificate = async (template, empresa, usuarios) => {
    const promises = usuarios.map(async (usuario) => {
        const doc = new Docxtemplater(template);
        doc.setData({
            empresa: empresa,
            nome: usuario.nome,
            cpf: usuario.cpf
        });
        doc.render();
        const buffer = doc.getZip().generate({ type: 'nodebuffer' });
        // Salva o certificado gerado com o nome do usuário
        await promisify(fs.writeFile)(`${usuario.nome}_certificado.docx`, buffer);
    });
    await Promise.all(promises);
};

// Rota para gerar certificados
app.post('/gerar-certificados', async (req, res) => {
    try {
        // Carrega o modelo do certificado
        const template = await loadTemplate();

        // Dados da empresa e dos usuários (exemplo)
        const empresa = req.body.empresa;
        const usuarios = req.body.usuarios;

        // Gera os certificados
        await generateCertificate(template, empresa, usuarios);

        res.send('Certificados gerados com sucesso!');
    } catch (error) {
        console.error('Erro ao gerar certificados:', error);
        res.status(500).send('Erro ao gerar certificados');
    }
});

// Inicia o servidor
app.listen(port, () => {
    console.log(`Servidor CertifEasy rodando em http://localhost:${port}`);
});