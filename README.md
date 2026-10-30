# E-Cíclica • MRS Logística

Plataforma para **preencher a cíclica (plano de manutenção)** pelo celular ou computador — no estilo do app do AppSheet, com a mesma identidade visual do Portal de Pátios.
HTML + CSS + JavaScript puro, **sem build**. Roda no **GitHub Pages**.

## O que tem
- **Login com usuário e senha**, dois perfis:
  - **Administrador** — envia/troca o documento do mês (.xlsx), cadastra funcionários, redefine senhas, edita equipes, vê tudo e exporta.
  - **Funcionário** — preenche as OS e exporta a planilha preenchida.
- **Lista das OS** agrupada por pátio → ativo, com pesquisa (OS, ativo, serviço, colaborador) e filtros por pátio, situação (sem preencher / executadas / pendentes) e aba.
- **Preenchimento** de uma OS: OK (Executado/Pendente), Dia, H início, H fim (botão “agora”), Manutenção imediata (S/N), Necessidade de manutenção (S/N), Colaboradores (com sugestões) e Equipe.
- **Preencher várias OS de uma vez**: toque em **Selecionar** (ou segure o dedo numa OS), marque as OS — tocar no cabeçalho de um ativo marca todas dele, e “Todas da lista” marca tudo o que está filtrado — e toque **Preencher**.
- **Painel**: andamento geral, por pátio, por colaborador e últimos preenchimentos.
- **Exportar**: baixa o **mesmo arquivo .xlsx enviado**, com as colunas preenchidas (mantém abas, cores, filtros, larguras). As abas de cada pátio (FCA, FEA, …) também recebem os preenchimentos. Opcional: coluna **EQUIPE** no final.
- Funciona **offline** depois do primeiro acesso (com Firebase): o que for preenchido sem sinal sobe sozinho quando a internet volta.

## Como o sistema lê a planilha
Na aba Admin → **Documento do mês**, envie o .xlsx. O sistema procura em cada aba a linha de cabeçalho com as colunas **OS**, **OK** e **DIA** (e também PATIO, ATIVO, SERVIÇO, Descrição, H INICIO, H FIM, Manutenção imediata, Há necessidade, COLABORADORES — em qualquer ordem).
- Abas com OS próprias (ex.: `01° DIA UTIL`, `5° DIA UTIL`) = **preencher**.
- Abas que só repetem OS de outra (ex.: uma aba por pátio) = **espelho** (recebem os mesmos preenchimentos na exportação).
- Abas ocultas são ignoradas.

Todo mês: envie a planilha nova e marque “Tornar este o documento ativo”. Os meses anteriores continuam na lista (dá para baixar preenchida ou excluir).

---

## Configurar o Firebase (necessário para a equipe usar junto) — grátis, ~10 min
Sem isso o site roda em **modo local (demonstração)**: dados só no navegador em uso.

1. Acesse https://console.firebase.google.com → **Adicionar projeto** (pode desativar o Google Analytics).
2. **Authentication** → Começar → aba **Método de login** → ative **E-mail/senha**.
3. **Firestore Database** → Criar banco de dados → modo **produção** → região `southamerica-east1 (São Paulo)`.
4. Ainda no Firestore → aba **Regras** → apague tudo, cole o conteúdo do arquivo [`firestore.rules`](firestore.rules) → **Publicar**.
5. ⚙️ Configurações do projeto → **Seus apps** → ícone **Web `</>`** → registre o app (não precisa de Hosting). Copie os valores do `firebaseConfig`.
6. No GitHub, abra o arquivo [`config.js`](config.js) → lápis ✏️ → cole `apiKey`, `authDomain`, `projectId`, `storageBucket`, `messagingSenderId`, `appId` → **Commit**.
7. Authentication → **Configurações** → **Domínios autorizados** → adicione `matheusprado0096.github.io`.
8. Abra o site: aparece **Primeiro acesso** → crie o seu usuário **administrador** (faça isso logo — o primeiro a criar vira o admin).
9. Admin → Usuários → cadastre os funcionários (usuário + senha inicial + equipe). Admin → Documento do mês → envie a planilha.

> A `apiKey` do Firebase não é segredo (ela vai para o navegador de qualquer forma); quem protege os dados são as **regras** do passo 4: só usuários cadastrados e ativos leem/preenchem, e só o admin troca documento e usuários. A planilha **não** fica no repositório público — fica no Firestore, protegida pelo login.

### Senhas
- O funcionário troca a própria senha em: botão com as iniciais (canto superior) → **Trocar minha senha**.
- Esqueceu? O admin usa **Nova senha** em Admin → Usuários (o login continua o mesmo).
- **Desativar** bloqueia o acesso sem apagar o histórico.

### Limites do plano grátis (Spark)
50 mil leituras e 20 mil gravações por dia — sobra para uma equipe preenchendo ~1.500 OS por mês (os preenchimentos ficam agrupados por pátio, então abrir a lista custa poucas leituras).

## Publicar no GitHub Pages
**Settings → Pages → Build and deployment → Deploy from a branch → `main` / `(root)`** → o site fica em `https://matheusprado0096.github.io/E-Ciclica/`.

Teste no computador: `python3 -m http.server 8000` e abra http://localhost:8000.

## Estrutura
```
index.html                 página única
config.js                  ⚙️ dados do Firebase e equipes padrão
firestore.rules            regras de segurança (colar no Firebase)
assets/css/style.css       estilos (identidade MRS, tema claro/escuro)
assets/js/app.js           interface: login, lista, preenchimento, lote, painel, exportação, admin
assets/js/xlsx.js          leitura da planilha e geração do .xlsx preenchido
assets/js/store-firebase.js  dados no Firebase (equipe)
assets/js/store-local.js     dados no navegador (modo demonstração)
assets/vendor/jszip.min.js   JSZip 3.10.1 (MIT)
sw.js                      funcionamento offline
```
