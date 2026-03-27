# Súmulas COIC/CBIC

Gerador de súmulas de reunião para a Comissão de Obras Industriais e Corporativas da CBIC.

## Deploy no VPS

```bash
# 1. Clone o repositório
git clone https://github.com/SEU_USUARIO/sumulas-coic.git /opt/sumulas-coic
cd /opt/sumulas-coic

# 2. Instale dependências
npm install --production

# 3. Inicie com PM2
pm2 start src/server.js --name sumulas-coic --max-memory-restart 300M
pm2 save
pm2 startup

# 4. Acesse
# http://SEU_IP:3000
```

## Atualizar

```bash
cd /opt/sumulas-coic
git pull
pm2 restart sumulas-coic
```

## Importante

O arquivo `templates/Sumula_Template_v2.docx` precisa estar presente para gerar os documentos Word.
