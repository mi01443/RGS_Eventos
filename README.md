# 🥂 Jantar dos Casais — Guia de Configuração

## Arquivos incluídos
- `index.html` → app completo (hospede no GitHub Pages)
- `Code.gs` → backend Google Apps Script (cole no Google Sheets)

---

## Passo 1 — Configurar o Google Sheets

1. Crie uma planilha nova em [sheets.google.com](https://sheets.google.com)
2. No menu, vá em **Extensões → Apps Script**
3. Apague o código padrão e cole todo o conteúdo do arquivo `Code.gs`
4. Salve (Ctrl+S) e clique em **Executar → setup** (isso cria as abas automaticamente)
5. Conceda as permissões solicitadas

---

## Passo 2 — Publicar o Apps Script como Web App

1. No Apps Script, clique em **Implantar → Nova implantação**
2. Tipo: **Aplicativo da Web**
3. Configurações:
   - **Executar como:** Eu mesmo
   - **Quem tem acesso:** Qualquer pessoa
4. Clique em **Implantar** e copie a **URL do Web App** gerada

---

## Passo 3 — Configurar o index.html

Abra `index.html` e localize a linha:

```javascript
const SHEET_URL = 'https://script.google.com/macros/s/SEU_ID_AQUI/exec';
```

Substitua pelo URL copiado no passo anterior.

---

## Passo 4 — Hospedar no GitHub Pages

1. Crie um repositório no GitHub (pode ser privado ou público)
2. Faça upload do `index.html` (somente esse arquivo)
3. Vá em **Settings → Pages**
4. Em "Source", selecione **Deploy from a branch → main → / (root)**
5. Aguarde ~1 minuto e acesse a URL gerada (ex: `https://seuusuario.github.io/jantar-casais`)

---

## Passo 5 — Configurar o evento no app

1. Acesse o app pelo GitHub Pages
2. Clique na aba **Admin**
3. Senha padrão: `admin123` (troque na aba Evento)
4. Preencha:
   - **Aba Evento:** título, data, horário, local, valor, WhatsApp do admin
   - **Aba PIX:** chave PIX, nome do recebedor, banco
   - **Aba Lista:** adicione os itens de comida (doces e salgados)

---

## Funcionalidades do app

### Área pública
- ✅ Visualizar evento (data, local, valor)
- ✅ Ver lista de itens disponíveis/reservados/confirmados
- ✅ Reservar um item (com nome e WhatsApp)
- ✅ Ver QR Code do PIX + copiar chave
- ✅ Enviar comprovante via WhatsApp para o admin

### Painel Admin
- ✅ Autenticação por senha
- ✅ Cadastrar/editar evento
- ✅ Configurar dados do PIX
- ✅ Montar lista de itens (adicionar/remover)
- ✅ Ver pagamentos pendentes e confirmados
- ✅ Confirmar pagamento → abre WhatsApp automaticamente com mensagem de confirmação

---

## Observações

- O QR Code do PIX exibe a **chave PIX** formatada — para um QR Code de pagamento completo (com valor embutido), seria necessário implementar o padrão EMV do Banco Central. Para simplificar, o participante usa a chave manualmente ou via QR Code básico.
- O app funciona **offline-first** com dados em memória enquanto o Sheets não está configurado.
- Todos os dados ficam nas abas: `Config`, `Foods` e `Reservations` da sua planilha.
