<div align="center">
<img width="1200" height="475" alt="GHBanner" src="https://github.com/user-attachments/assets/0aa67016-6eaf-458a-adb2-6e31a0763ed6" />
</div>

# Run and deploy your AI Studio app

This contains everything you need to run your app locally.

View your app in AI Studio: https://ai.studio/apps/8617308d-c341-459d-9b04-0acf54f18b2c

## Run Locally

**Prerequisites:**  Node.js


1. Install dependencies:
   `npm install`
2. Set the `GEMINI_API_KEY` in [.env.local](.env.local) to your Gemini API key
3. Run the app:
   `npm run dev`

## Sincronização entre dispositivos

O navegador busca atualizações a cada 30 segundos enquanto a página está visível,
ao recuperar a conexão e ao voltar para o app. Alterações locais e a última base
confirmada ficam salvas juntas no navegador; envios com erro continuam pendentes.
A confirmação de envio só aparece depois da resposta de sucesso da API.

A API recebe alterações por ID e campo. Registros novos são acrescentados com
`values.append`; registros existentes são atualizados nas próprias células.
Exclusões usam `syncDeleted` (e `DELETADO` nas solicitações). A sincronização comum
não limpa nem substitui as tabelas inteiras. Se duas pessoas alterarem o mesmo
campo do mesmo registro, prevalece a última gravação; não há resolução interativa
de conflitos. As abas derivadas de relatório continuam sendo atualizadas depois
da gravação; essa atualização não é uma transação com os registros principais.

### Validação

- `npm test`: regressões de concorrência, reenvio, exclusão, recuperação local e merge.
- `npm run lint`: verificação TypeScript.
- `npm run build`: compilação de produção.
- Em dois navegadores separados, lançar em cada um, voltar ao outro e verificar os
  dois IDs. Repetir após ficar sem internet e recarregar o navegador com envio pendente.

### Publicação

Publicar frontend e API juntos. A API rejeita o formato antigo de gravação integral;
aberturas antigas devem ser atualizadas nos celulares e computadores. Não limpar
o armazenamento do navegador quando houver lançamentos pendentes. Não há migração
manual: as colunas necessárias são adicionadas às abas existentes pela API.
As credenciais da conta de serviço e o acesso de Editor à mesma planilha continuam
necessários. O endereço da planilha deve ser igual em todos os aparelhos.
