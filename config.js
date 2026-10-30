// ===== E-Cíclica • configuração =====
// 1) Crie um projeto grátis no Firebase (veja o README, passo a passo).
// 2) Cole aqui os dados do "firebaseConfig" do seu app Web.
// Enquanto apiKey estiver vazio, o sistema roda em MODO LOCAL (demonstração):
// os dados ficam só no aparelho/navegador em uso.
window.ECICLICA_CONFIG = {
  firebase: {
    apiKey: "",
    authDomain: "",
    projectId: "",
    storageBucket: "",
    messagingSenderId: "",
    appId: ""
  },

  // domínio "de mentira" usado para transformar o usuário em e-mail no Firebase
  // (ex.: matheus -> matheus-xxxx@eciclica.app). Não precisa existir.
  dominioLogin: "eciclica.app",

  // equipes que aparecem como botões no preenchimento (o admin pode mudar no sistema)
  equipes: ["2W", "2X", "2Y", "2Z", "2R"]
};
