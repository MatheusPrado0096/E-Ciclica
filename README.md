# E-Cíclica • MRS Logística

Plataforma para **preencher a cíclica (plano de manutenção)** pelo celular ou computador — no estilo do app do AppSheet, com a mesma identidade visual do Portal de Pátios.
HTML + CSS + JavaScript puro, **sem build**. Roda no **GitHub Pages**.

## O que tem
- **Entrada por matrícula** — todos têm o mesmo acesso. No primeiro acesso a pessoa informa **matrícula + nome** (e, se quiser, a equipe); depois entra só com a matrícula. Nome e equipe podem ser alterados em “Meus dados” (botão com as iniciais).
- **Senha de upload** — enviar, ativar ou excluir a planilha do mês e editar as equipes só funciona depois de digitar a senha de upload na aba **Documento** (fica liberado até sair/fechar o navegador). Com Firebase, a senha é conferida pelo servidor (nas regras), não só pela tela.
- **Lista das OS** agrupada por pátio → ativo, com pesquisa (OS, ativo, serviço, colaborador), **seletor de pátio com pesquisa** (digite “FP” e veja FPA, FPI… com o andamento de cada um), filtros por situação e aba.
- **Preenchimento** de uma OS: OK (Executado/Pendente), Dia, H início, H fim (botão “agora”), Manutenção imediata (S/N), Necessidade de manutenção (S/N), Colaboradores (com sugestões) e Equipe.
- **Preencher várias OS de uma vez**: toque em **Selecionar** (ou segure o dedo numa OS), marque as OS — tocar no cabeçalho de um ativo marca todas dele, e “Todas da lista” marca tudo o que está filtrado — e toque **Preencher**.
- **Painel de monitoramento** com filtros de **período** (hoje, 7 dias, 30 dias ou datas personalizadas), **pátio**, **equipe** e **colaborador**, separado em visões: Visão geral (números + preenchimentos por dia), Pátios, Equipes, Colaboradores, Atenção (OS com manutenção imediata/necessidade) e Atividade (últimos preenchimentos). No celular os filtros ficam recolhidos num botão.
- **Exportar**: baixa o **mesmo arquivo .xlsx enviado**, com as colunas preenchidas (mantém abas, cores, filtros, larguras). As abas de cada pátio (FCA, FEA, …) também recebem os preenchimentos. Opcional: coluna **EQUIPE** no final.
- Funciona **offline** depois do primeiro acesso (com Firebase): o que for preenchido sem sinal sobe sozinho quando a internet volta.

## Como o sistema lê a planilha
Na aba **Documento**, digite a senha de upload e envie o .xlsx. O sistema procura em cada aba a linha de cabeçalho com as colunas **OS**, **OK** e **DIA** (e também PATIO, ATIVO, SERVIÇO, Descrição, H INICIO, H FIM, Manutenção imediata, Há necessidade, COLABORADORES — em qualquer ordem).
- Abas com OS próprias (ex.: `01° DIA UTIL`, `5° DIA UTIL`) = **preencher**.
- Abas que só repetem OS de outra (ex.: uma aba por pátio) = **espelho** (recebem os mesmos preenchimentos na exportação).
- Abas ocultas são ignoradas.

Todo mês: envie a planilha nova e marque “Tornar este o documento ativo”. Os meses anteriores continuam na lista (dá para baixar preenchida ou excluir).

---

## Configurar o Firebase (necessário para a equipe usar junto) — grátis, ~10 min
Sem isso o site roda em **modo local (demonstração)**: dados só no navegador em uso.

1. Acesse https://console.firebase.google.com → **Adicionar projeto** (pode desativar o Google Analytics).
2. **Authentication** → Começar → aba **Método de login** → ative **Anônimo**.
3. **Firestore Database** → Criar banco de dados → modo **produção** → região `southamerica-east1 (São Paulo)`.
4. Ainda no Firestore → aba **Regras** → apague tudo, cole o conteúdo do arquivo [`firestore.rules`](firestore.rules) → **Publicar**.
5. ⚙️ Configurações do projeto → **Seus apps** → ícone **Web `</>`** → registre o app (não precisa de Hosting). Copie os valores do `firebaseConfig`.
6. No GitHub, abra o arquivo [`config.js`](config.js) → lápis ✏️ → cole `apiKey`, `authDomain`, `projectId`, `storageBucket`, `messagingSenderId`, `appId` → **Commit**.
7. Authentication → **Configurações** → **Domínios autorizados** → adicione `matheusprado0096.github.io`.
8. Abra o site, entre com sua matrícula, vá em **Documento**, digite a senha de upload e envie a planilha do mês.

> A `apiKey` do Firebase não é segredo (ela vai para o navegador de qualquer forma); quem protege os dados são as **regras** do passo 4. A planilha **não** fica no repositório público — fica no Firestore.
> A entrada é só por matrícula (sem senha pessoal), como pedido: qualquer pessoa com o link que digitar uma matrícula cadastrada entra com aquele nome. Para o que importa (trocar/excluir documento) existe a senha de upload.

### Trocar a senha de upload
A senha não aparece no código, só o “hash” dela. Para trocar:
1. Gere os dois formatos do hash da nova senha (no navegador, F12 → Console):
   ```js
   const h = new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode('NOVA-SENHA')));
   console.log('hex:', [...h].map(b => b.toString(16).padStart(2, '0')).join(''), '\nbase64:', btoa(String.fromCharCode(...h)));
   ```
2. Coloque o **hex** em `firestore.rules` (função `senhaOk`) e publique as regras de novo no Firebase.
3. Coloque o **base64** em `assets/js/store-local.js` (`UPLOAD_HASH`).

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
assets/js/app.js           interface: login, lista, preenchimento, lote, painel, exportação, documento
assets/js/xlsx.js          leitura da planilha e geração do .xlsx preenchido
assets/js/store-firebase.js  dados no Firebase (equipe)
assets/js/store-local.js     dados no navegador (modo demonstração)
assets/vendor/jszip.min.js   JSZip 3.10.1 (MIT)
sw.js                      funcionamento offline
```
