# Portal de Relatórios (Interview & Senior)

Este é um projeto web moderno, seguro e responsivo projetado para automatizar o download de relatórios do portal **Interview/Reportview**.

Ele substitui scripts locais em Python/VBScript por uma interface web premium, protegida por senha, que pode ser acessada de qualquer lugar (computador, celular ou tablet) e compartilhada de forma segura com outros membros da equipe.

---

## 🚀 Funcionalidades

1. **Acesso Restrito ao Portal**: Login protegido por usuário e senha (`gestão` / `gestão`) para garantir que apenas pessoas autorizadas acessem a ferramenta.
2. **Campos Dinâmicos para Credenciais**: Cada usuário pode inserir seu próprio usuário e senha do Interview (sem senhas salvas no código fonte!).
3. **Salvamento Seguro Local (Opcional)**: As credenciais podem ser salvas no navegador (`localStorage`) para que você não precise redigitá-las a cada acesso.
4. **Seletor de Relatórios**: Baixe o relatório de colaboradores padrão (ID 132) ou digite qualquer outro ID de relatório do Interview de forma manual.
5. **Console de Status em Tempo Real**: Veja o passo a passo da integração, desde a autenticação até a extração dos dados, em um terminal de logs interativo.
6. **Pré-visualização Interativa**: Visualize e faça buscas na amostra dos dados (primeiras 10 linhas) diretamente na tela antes de exportar.
7. **Conversão Instantânea para Excel (.xlsx)**: Criação do arquivo Excel no próprio navegador utilizando a biblioteca SheetJS (processamento super rápido).
8. **Proxy de Rede Seguro**: O backend do Vercel atua como um túnel seguro para realizar a autenticação e o download dos dados do servidor de relatórios, superando restrições de CORS e Mixed Content dos navegadores modernos.

---

## 🛠️ Estrutura do Projeto

* `index.html`: Estrutura HTML da página única (SPA) contendo a tela de Login e a tela de Dashboard.
* `style.css`: Estilização premium baseada em Glassmorphism, com cores sofisticadas, sombras suaves, animações e responsividade.
* `script.js`: Toda a lógica da aplicação (validação de acesso, manipulação de logs, renderização de tabelas e geração do Excel).
* `vercel.json`: Arquivo de configuração de rotas e URLs limpas para o Vercel.
* `api/baixar.js`: Função Serverless em Node.js que realiza a conexão HTTP segura com o servidor de relatórios utilizando variáveis de ambiente.

---

## 📥 Como Publicar no GitHub e Implantar no Vercel

Siga o passo a passo abaixo para colocar o seu portal online de graça:

### Passo 1: Criar um Repositório no GitHub
1. Acesse o [GitHub](https://github.com/) e faça login.
2. Clique no botão **"New"** (Novo repositório) no canto superior esquerdo.
3. Dê um nome ao seu repositório (ex: `relatorios-interview-senior`).
4. Deixe o repositório como **Private** (Privado) se preferir que apenas você e quem você convidar possam ver o código.
5. Clique em **"Create repository"** (Criar repositório).

### Passo 2: Subir os arquivos para o GitHub
Você pode fazer isso direto pelo site do GitHub:
1. Na página do repositório recém-criado, clique no link **"uploading an existing file"** (enviar um arquivo existente).
2. Arraste e solte os seguintes arquivos e pastas do seu computador para o navegador:
   * `index.html`
   * `style.css`
   * `script.js`
   * `vercel.json`
   * Pasta `api` (com o arquivo `baixar.js` dentro)
3. Clique em **"Commit changes"** (Salvar alterações) no final da página.

---

### Passo 3: Publicar no Vercel e Configurar a Variável de Ambiente
1. Acesse o site do [Vercel](https://vercel.com/) e faça login usando sua conta do **GitHub**.
2. No painel principal da Vercel, clique no botão **"Add New..."** e selecione **"Project"** (Projeto).
3. Na lista de repositórios do GitHub, clique em **"Import"** ao lado do repositório que você criou.
4. Na tela de configurações do projeto, expanda a seção **"Environment Variables"** (Variáveis de Ambiente).
5. Adicione a variável que define para onde as requisições de relatório devem ser enviadas:
   * **Key (Chave)**: `REPORT_SYSTEM_DOMAIN`
   * **Value (Valor)**: `reportview.cocal.com.br` (ou o domínio do seu sistema de relatórios)
6. Clique em **"Add"** (Adicionar).
7. Clique no botão **"Deploy"** (Implantar).
8. Pronto! A Vercel gerará um link público seguro (ex: `https://seu-projeto.vercel.app`) para você acessar!

---

## 🔒 Segurança de Credenciais
* **Variável de Ambiente**: Ao utilizar a variável `REPORT_SYSTEM_DOMAIN` no Vercel, o domínio da sua empresa **nunca** fica visível no código público do GitHub, mantendo o sistema 100% anônimo para terceiros.
* **Segurança no Servidor**: A senha do seu Interview **nunca** fica gravada em nenhum servidor ou banco de dados externo. O Vercel apenas faz a ponte (como se fosse um túnel temporário) e envia o dado diretamente para o servidor de relatórios.
* **Segurança no Navegador**: Se você selecionar a opção "Salvar credenciais localmente", o usuário e a senha serão armazenados na memória local do seu navegador (`localStorage`). Ninguém na internet terá acesso a esse dado além de você.
