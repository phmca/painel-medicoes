const firebaseConfig = {
  apiKey: "AIzaSyDJgTjFtq7WXPZlKmMQ_c6H8vVoZTI4JaM",
  authDomain: "medicoes-dnit.firebaseapp.com",
  databaseURL: "https://medicoes-dnit-default-rtdb.firebaseio.com",
  projectId: "medicoes-dnit",
  storageBucket: "medicoes-dnit.firebasestorage.app",
  messagingSenderId: "458189906562",
  appId: "1:458189906562:web:4149641ef4dfcc7786ca9b"
};

firebase.initializeApp(firebaseConfig);
const database = firebase.database();
const auth = firebase.auth();

const REF_CONTRATOS = 'execucao_financeira/contratos';
const REF_CONFIG = 'execucao_financeira/config';
const REF_USUARIOS = 'usuarios';
const REF_MEDICOES = 'execucao_financeira/medicoes_sei';

let MESES = [
    '2024-12', '2025-01', '2025-02', '2025-03', '2025-04', '2025-05',
    '2025-06', '2025-07', '2025-08', '2025-09', '2025-10', '2025-11',
    '2025-12', '2026-01', '2026-02', '2026-03', '2026-04', '2026-05',
    '2026-06', '2026-07', '2026-08', '2026-09'
];

const XLSX_CONFIG = {
    aba: 'Execução Financeira',
    linhaInicio: 11, linhaFim: 58, colunaContrato: 1, colunaUL: 0, colunaBR: 2,
    colunaServico: 3, colunaContratada: 5, linhaCabecalho: 8, colunaInicioFixa: 28,
    colunasIgnorar: [29, 41, 42], maxColunasVazias: 3, colunaVigencia: 8,
    colunaUltimaMedLancada: 10, colunaMesUltimaMedLancada: 11, colunaUltimaMedPaga: 12,
    colunaMesUltimaMedPaga: 13, colunaSaldoEmpenho: 18
};

const XLSX_CONFIG_MEDICOES = {
    aba: 'Medições SEI', linhaInicio: 2, colContrato: 0,
    colNumMedicao: 1, colProcPagamento: 2, colProcSuporte: 3
};

const DADOS_BASE = [
    { id: 'c_397_2024', ul: 'SOBRAL', contrato: '397/2024', br: '222', servico: 'Manutenção - PATO', contratada: 'Insttale Engenharia Ltda', situacao: 'Ativo', vigencia: '', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 12142259.20, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_347_2024', ul: 'SOBRAL', contrato: '347/2024', br: '222;403', servico: 'Manutenção - PATO', contratada: 'COSAMPA Projetos e Construções', situacao: 'Ativo', vigencia: '', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 10000000, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_404_2023', ul: 'SOBRAL', contrato: '404/2023', br: '222,402,403', servico: 'BR Legal', contratada: 'Consórcio Ceará Via Segura', situacao: 'Ativo', vigencia: '01/12/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 500000, loa: 3964408.90, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_250_2026', ul: 'FORTALEZA', contrato: '250/2026', br: '402;403', servico: 'Manutenção - PATO', contratada: 'Vanguarda Construções e Serviçõs de Conservação Viária LTDA', situacao: 'Ativo', vigencia: '29/09/2029', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_175_2025', ul: 'ICÓ', contrato: '175/2025', br: '116;226', servico: 'Manutenção - PATO', contratada: 'LCM Construção e Comércio S/A', situacao: 'Ativo', vigencia: '07/09/2027', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_236_2025', ul: 'ICÓ', contrato: '236/2025', br: '116', servico: 'Manutenção - PATO', contratada: 'LCM Construção e Comércio S/A', situacao: 'Concluído', vigencia: '05/10/2027', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_235_2025', ul: 'ICÓ', contrato: '235/2025', br: '230', servico: 'Manutenção - PATO', contratada: 'Insttale Engenharia Ltda', situacao: 'Ativo', vigencia: '06/10/2027', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_039_2026', ul: 'ICÓ', contrato: '039/2026', br: '116', servico: 'Obra de Emergência', contratada: 'FFX Engenharia e Serviços LTDA', situacao: 'Ativo', vigencia: '24/11/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_424_2023', ul: 'ICÓ', contrato: '424/2023', br: '116;226;230', servico: 'BR Legal', contratada: 'ESSE Eng. Sinal. Serv. Especiais Ltda', situacao: 'Ativo', vigencia: '01/12/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 500000, loa: 3230000, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_688_2022', ul: 'FORTALEZA', contrato: '688/2022', br: '116', servico: 'Manutenção - PATO', contratada: 'Insttale Engenharia Ltda', situacao: 'Ativo', vigencia: '04/07/2024', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 10436917.37, loa: 15000000, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_505_2024', ul: 'FORTALEZA', contrato: '505/2024', br: '020', servico: 'Manutenção - PATO', contratada: 'ESSE Eng. Sinal. Serv. Especiais Ltda', situacao: 'Aguardando', vigencia: '', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 5000000, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_092_2024', ul: 'FORTALEZA', contrato: '092/2024', br: '222', servico: 'Manutenção - PATO', contratada: 'Insttale Engenharia Ltda', situacao: 'Ativo', vigencia: '', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 17400000, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_124_2026', ul: 'FORTALEZA', contrato: '124/2026', br: '122', servico: '', contratada: 'Insttale Engenharia Ltda', situacao: 'Ativo', vigencia: '25/08/2029', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_240_2024', ul: 'FORTALEZA', contrato: '240/2024', br: '122', servico: 'Manutenção - PATO', contratada: 'LCM Construção e Comércio S/A', situacao: 'Ativo', vigencia: '23/09/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_372_2023', ul: 'FORTALEZA', contrato: '372/2023', br: '116;020;222', servico: 'BR Legal', contratada: 'ESSE Eng. Sinal. Serv. Especiais Ltda', situacao: 'Ativo', vigencia: '01/12/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 500000, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_542_2026', ul: 'FORTALEZA', contrato: '542/2026', br: '116', servico: 'Manutenção - PATO', contratada: 'CCL - Construtora Capital LTDA', situacao: 'Cadastrado', vigencia: '20/01/2030', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_515_2022', ul: 'RUSSAS', contrato: '515/2022', br: '116', servico: 'Manutenção - PATO', contratada: 'Insttale Engenharia Ltda', situacao: 'Ativo', vigencia: '27/09/2027', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_385_2021', ul: 'RUSSAS', contrato: '385/2021', br: '116;437', servico: 'Manutenção - PATO', contratada: 'LCM Construção e Comércio S/A', situacao: 'Ativo', vigencia: '09/08/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_413_2021', ul: 'RUSSAS', contrato: '413/2021', br: '304', servico: 'Manutenção - PATO', contratada: 'Insttale Engenharia Ltda', situacao: 'Ativo', vigencia: '11/08/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_040_2026', ul: 'RUSSAS', contrato: '040/2026', br: '116;122', servico: 'Obra de Emergência', contratada: 'Edro Engenharia LTDA', situacao: 'Ativo', vigencia: '25/11/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_041_2026', ul: 'RUSSAS', contrato: '041/2026', br: '116', servico: 'Obra de Emergência', contratada: 'Edro Engenharia LTDA', situacao: 'Ativo', vigencia: '10/11/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_421_2023', ul: 'RUSSAS', contrato: '421/2023', br: '116;304;122;437', servico: 'BR Legal', contratada: 'ESSE Eng. Sinal. Serv. Especiais Ltda', situacao: 'Ativo', vigencia: '01/12/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 500000, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_331_2026', ul: 'RUSSAS', contrato: '331/2026', br: '116;304', servico: 'Manutenção - PATO', contratada: 'Insttale Engenharia Ltda', situacao: 'Ativo', vigencia: '19/11/2029', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_508_2024', ul: 'BOA VIAGEM', contrato: '508/2024', br: '226;404', servico: 'Manutenção - PATO', contratada: 'AL Teixeira Pinheira', situacao: 'Ativo', vigencia: '07/01/2027', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_408_2024', ul: 'BOA VIAGEM', contrato: '408/2024', br: '226', servico: 'Manutenção - PATO', contratada: 'CSR - Construções e Serviços Rodoviários Ltda', situacao: 'Ativo', vigencia: '05/12/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_866_2024', ul: 'BOA VIAGEM', contrato: '866/2024', br: '020', servico: 'Manutenção - PATO', contratada: 'CivilPav Construções', situacao: 'Ativo', vigencia: '16/05/2027', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_634_2024', ul: 'BOA VIAGEM', contrato: '634/2024', br: '020', servico: 'Manutenção - PATO', contratada: 'CSR - Construções e Serviços Rodoviários Ltda', situacao: 'Ativo', vigencia: '25/02/2027', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_666_2025', ul: 'BOA VIAGEM', contrato: '666/2025', br: '020', servico: '', contratada: 'CIVILPAV CONSTRUÇOES LTDA', situacao: 'Ativo', vigencia: '31/03/2029', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_043_2026', ul: 'BOA VIAGEM', contrato: '043/2026', br: '226', servico: 'Obra de Emergência', contratada: 'Edro Engenharia LTDA', situacao: 'Ativo', vigencia: '14/11/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_420_2023', ul: 'BOA VIAGEM', contrato: '420/2023', br: '020;226;405', servico: 'BR Legal', contratada: 'Consórcio Sinal CE', situacao: 'Ativo', vigencia: '01/12/2026', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 500000, loa: 0, remanejMenos: 0, remanejMais: 0 },
    { id: 'c_189_2023', ul: 'SMT', contrato: '189/2023', br: 'TODA MALHA', servico: 'Manutenção - Supervisão', contratada: 'Maia Melo', situacao: 'Ativo', vigencia: '27/02/2027', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0 }
];

let dados = [], dadosFiltrados = [], modoEdicao = false, dadosCarregados = false, salvando = false, debounceTimers = {};
let arquivoAtual = null, dadosPreview = null, ulAtivaFiltro = 'todas', selecionadosIds = new Set();
let loginManualFeito = false, loginEmProgresso = false, appAbertoPara = null;
let medicoesSEI = [], medicoesFiltradas = [], contratosExpandidos = new Set(), idMedicaoEmEdicao = null;

function limparNomeContrato(nome) {
    if (!nome) return '';
    return String(nome).toUpperCase().replace(/\s+/g, '').trim();
}

async function fazerLogin(event) {
    if (event) event.preventDefault();
    const btn = document.getElementById('btnLogin');
    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password_fake').value; 
    const errorEl = document.getElementById('loginError');

    if (btn.disabled) return false;

    errorEl.classList.remove('show');
    if (!username || !password) {
        errorEl.textContent = '❌ Preencha usuário e senha!';
        errorEl.classList.add('show');
        return false;
    }

    document.getElementById('password').value = password;
    btn.disabled = true;
    btn.textContent = '⏳ Entrando...';
    loginEmProgresso = true;

    try {
        await auth.signOut();
        loginManualFeito = true;
        const email = username.includes('@') ? username : username + '@interno.local';
        const cred = await auth.signInWithEmailAndPassword(email, password);
        await abrirAppParaUsuario(cred.user);
        return false;
    } catch (erro) {
        loginManualFeito = false;
        let msg = '❌ Usuário ou senha incorretos!';
        if (erro.code === 'auth/invalid-login-credentials') {
            msg = '❌ Usuário ou senha incorretos. Verifique o cadastro no Firebase.';
        }
        errorEl.textContent = msg;
        errorEl.classList.add('show');
        return false;
    } finally {
        loginEmProgresso = false;
        btn.disabled = false;
        btn.textContent = '🔐 Entrar';
    }
}

function toggleSenha(btn) {
    const inputFake = document.getElementById('password_fake');
    const senhaOculta = inputFake.style.webkitTextSecurity !== 'none';
    if (senhaOculta) {
        inputFake.style.webkitTextSecurity = 'none';
        btn.textContent = '🙈';
    } else {
        inputFake.style.webkitTextSecurity = 'disc';
        btn.textContent = '👁️';
    }
}

function fazerLogout() {
    if (!confirm('Deseja realmente sair?')) return;
    loginManualFeito = false;
    loginEmProgresso = false;
    appAbertoPara = null; 
    auth.signOut().then(() => {
        document.getElementById('appContainer').style.display = 'none';
        document.getElementById('modalTrocaSenha').style.display = 'none';
        document.getElementById('loginOverlay').style.display = 'flex';
        document.getElementById('username').value = '';
        document.getElementById('password_fake').value = '';
        document.getElementById('password_fake').style.webkitTextSecurity = 'disc';
        document.getElementById('password').value = '';
        dadosCarregados = false;
    });
}

async function salvarNovaSenha(event) {
    if (event) event.preventDefault();
    const novaSenha = document.getElementById('novaSenha').value;
    const confirmarSenha = document.getElementById('confirmarSenha').value;
    const erroEl = document.getElementById('erroTrocaSenha');

    erroEl.style.display = 'none';
    if (novaSenha.length < 6) {
        erroEl.textContent = '❌ A senha deve ter pelo menos 6 caracteres.';
        erroEl.style.display = 'block';
        return;
    }
    if (novaSenha !== confirmarSenha) {
        erroEl.textContent = '❌ As senhas não coincidem.';
        erroEl.style.display = 'block';
        return;
    }

    try {
        const user = auth.currentUser;
        if (!user) return;
        await user.updatePassword(novaSenha);
        const username = user.email.split('@')[0];
        await database.ref(`${REF_USUARIOS}/${username}`).update({ precisaTrocarSenha: false, dataTrocaSenha: new Date().toISOString() });
        alert('✅ Senha atualizada com sucesso! Você será redirecionado.');
        document.getElementById('modalTrocaSenha').style.display = 'none';
        document.getElementById('appContainer').style.display = 'block';
        if (!dadosCarregados) carregarDados();
    } catch (error) {
        erroEl.textContent = '❌ Erro ao atualizar senha: ' + error.message;
        erroEl.style.display = 'block';
    }
}

auth.onAuthStateChanged(async (user) => {
    if (loginEmProgresso) return;
    if (user && !loginManualFeito) {
        await auth.signOut();
        return;
    }
    if (user) {
        await abrirAppParaUsuario(user);
    } else {
        appAbertoPara = null;
        document.getElementById('appContainer').style.display = 'none';
        document.getElementById('modalTrocaSenha').style.display = 'none';
        document.getElementById('loginOverlay').style.display = 'flex';
        document.getElementById('usuarioNomeHeader').textContent = '-';
        dadosCarregados = false;
    }
});

async function abrirAppParaUsuario(user) {
    if (appAbertoPara === user.uid) return;
    appAbertoPara = user.uid;
    const username = user.email.split('@')[0];

    document.getElementById('loginOverlay').style.display = 'none';
    document.getElementById('appContainer').style.display = 'none';
    document.getElementById('modalTrocaSenha').style.display = 'none';

    try {
        if (username !== 'dev1') {
            const snapshot = await database.ref(`${REF_USUARIOS}/${username}`).once('value');
            const userData = snapshot.val();
            if (!userData) {
                await database.ref(`${REF_USUARIOS}/${username}`).set({ precisaTrocarSenha: true, primeiroAcesso: new Date().toISOString() });
            }
            if (!userData || userData.precisaTrocarSenha === true) {
                document.getElementById('modalTrocaSenha').style.display = 'flex';
                document.getElementById('usuarioNome').textContent = username;
                document.getElementById('usernameTroca').value = username;
                return;
            }
        }
        document.getElementById('appContainer').style.display = 'block';
        document.getElementById('usuarioNome').textContent = username;
        document.getElementById('usuarioNomeHeader').textContent = username;
        if (!dadosCarregados) carregarDados();
    } catch (error) {
        document.getElementById('appContainer').style.display = 'block';
        if (!dadosCarregados) carregarDados();
    }
}

function formatarMoeda(v) { return (!v && v !== 0) ? 'R$ 0,00' : Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); }
function formatarMoedaCurta(v) {
    if (!v && v !== 0) return 'R$ 0';
    if (v >= 1000000) return 'R$ ' + (v / 1000000).toFixed(2) + 'M';
    if (v >= 1000) return 'R$ ' + (v / 1000).toFixed(0) + 'K';
    return 'R$ ' + Number(v).toFixed(0);
}
function formatarNumeroInput(v) { return (!v && v !== 0) ? '0,00' : Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function parseNumero(str) {
    if (typeof str === 'number') return str;
    if (!str) return 0;
    let s = String(str).trim();
    if (s.includes(',') && s.includes('.')) s = s.replace(/\./g, '').replace(',', '.');
    else if (s.includes(',')) s = s.replace(',', '.');
    return parseFloat(s) || 0;
}
function getBadgeClass(s) {
    if (!s) return 'badge-cancelado';
    const l = s.toLowerCase();
    if (l.includes('ativo')) return 'badge-ativo';
    if (l.includes('concluído')) return 'badge-concluido';
    if (l.includes('aguardando')) return 'badge-aguardando';
    return 'badge-cancelado';
}
function mostrarStatus(msg, tipo) {
    const el = document.getElementById('saveStatus');
    el.textContent = msg; el.className = 'save-status';
    if (tipo === 'erro') el.classList.add('erro');
    else if (tipo === 'info') el.classList.add('info');
    el.style.display = 'block';
    clearTimeout(el._timeout);
    el._timeout = setTimeout(() => { el.style.display = 'none'; }, 3000);
}

function formatarDataBR(date) { return String(date.getDate()).padStart(2, '0') + '/' + String(date.getMonth() + 1).padStart(2, '0') + '/' + date.getFullYear(); }
function formatarMesAno(date) {
    const meses = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
    return `${meses[date.getMonth()]}/${String(date.getFullYear()).slice(2)}`;
}
function parseDataExcel(valor) {
    if (!valor && valor !== 0) return '';
    if (valor instanceof Date && !isNaN(valor)) return formatarDataBR(valor);
    if (typeof valor === 'number' && valor > 1 && valor < 100000) {
        try {
            const data = XLSX.SSF.parse_date_code(valor);
            if (data && data.y && data.m && data.d) return `${String(data.d).padStart(2, '0')}/${String(data.m).padStart(2, '0')}/${data.y}`;
        } catch (e) { }
    }
    const s = String(valor).trim();
    if (!s) return '';
    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return `${String(m[1]).padStart(2, '0')}/${String(m[2]).padStart(2, '0')}/${m[3]}`;
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return `${String(m[3]).padStart(2, '0')}/${String(m[2]).padStart(2, '0')}/${m[1]}`;
    m = s.match(/^([a-zç]{3,})[\-\/](\d{2,4})$/i);
    if (m) {
        const mapMeses = { jan: '01', fev: '02', mar: '03', abr: '04', mai: '05', jun: '06', jul: '07', ago: '08', set: '09', out: '10', nov: '11', dez: '12', janeiro: '01', fevereiro: '02', março: '03', abril: '04', maio: '05', junho: '06', julho: '07', agosto: '08', setembro: '09', outubro: '10', novembro: '11', dezembro: '12' };
        const nomeMes = m[1].toLowerCase();
        if (mapMeses[nomeMes]) return `${mapMeses[nomeMes]}/${m[2].length === 2 ? '20' + m[2] : m[2]}`;
    }
    return s;
}
function parseMesAnoExcel(valor) {
    if (!valor && valor !== 0) return '';
    if (valor instanceof Date && !isNaN(valor)) return formatarMesAno(valor);
    if (typeof valor === 'number' && valor > 1 && valor < 100000) {
        try {
            const data = XLSX.SSF.parse_date_code(valor);
            if (data && data.y && data.m) {
                const meses = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
                return `${meses[data.m - 1]}/${String(data.y).slice(2)}`;
            }
        } catch (e) { }
    }
    const s = String(valor).trim();
    if (!s) return '';
    let m = s.match(/^([a-zç]{3,})[\-\/](\d{2,4})$/i);
    if (m) {
        const mapMeses = { jan: '01', fev: '02', mar: '03', abr: '04', mai: '05', jun: '06', jul: '07', ago: '08', set: '09', out: '10', nov: '11', dez: '12' };
        const nomeMes = m[1].toLowerCase().substring(0, 3);
        if (mapMeses[nomeMes]) return `${nomeMes}/${m[2].length === 2 ? m[2] : m[2].slice(-2)}`;
    }
    m = s.match(/^\d{1,2}\/(\d{1,2})\/(\d{4})/);
    if (m) {
        const meses = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
        return `${meses[parseInt(m[1]) - 1]}/${m[2].slice(-2)}`;
    }
    return s;
}
function parseMesDoCabecalho(valor) {
    if (!valor) return null;
    if (valor instanceof Date && !isNaN(valor)) return valor.getFullYear() + '-' + String(valor.getMonth() + 1).padStart(2, '0');
    const s = String(valor).trim();
    if (!s) return null;
    let m = s.match(/^(\d{4})-(\d{1,2})(?:-\d{1,2})?/);
    if (m) {
        const mes = String(parseInt(m[2])).padStart(2, '0');
        if (parseInt(mes) >= 1 && parseInt(mes) <= 12) return m[1] + '-' + mes;
    }
    m = s.match(/^\d{1,2}\/(\d{1,2})\/(\d{4})/);
    if (m) {
        const mes = String(parseInt(m[1])).padStart(2, '0');
        if (parseInt(mes) >= 1 && parseInt(mes) <= 12) return m[2] + '-' + mes;
    }
    m = s.match(/([a-zç]{3,})[\/\s\-]+(\d{2,4})/i);
    if (m) {
        const mapMeses = { jan: '01', fev: '02', mar: '03', abr: '04', mai: '05', jun: '06', jul: '07', ago: '08', set: '09', out: '10', nov: '11', dez: '12', janeiro: '01', fevereiro: '02', março: '03', abril: '04', maio: '05', junho: '06', julho: '07', agosto: '08', setembro: '09', outubro: '10', novembro: '11', dezembro: '12' };
        const mes = mapMeses[m[1].toLowerCase()];
        if (mes) return `${m[2].length === 2 ? '20' + m[2] : m[2]}-${mes}`;
    }
    return null;
}
function detectarMesesDoCabecalho(matriz) {
    const linhaCab = matriz[XLSX_CONFIG.linhaCabecalho - 1];
    if (!linhaCab) return null;
    const mesesDetectados = []; let vazias = 0;
    for (let j = XLSX_CONFIG.colunaInicioFixa; j < linhaCab.length; j++) {
        if (XLSX_CONFIG.colunasIgnorar.includes(j)) continue;
        const mes = parseMesDoCabecalho(linhaCab[j]);
        if (mes) { mesesDetectados.push({ coluna: j, mes: mes }); vazias = 0; }
        else { vazias++; if (vazias >= XLSX_CONFIG.maxColunasVazias) break; }
    }
    return mesesDetectados.length > 0 ? mesesDetectados : null;
}
function gerarMesesPadrao() {
    const meses = []; let ano = 2024, mes = 12;
    for (let i = 0; i < 22; i++) {
        meses.push(ano + '-' + String(mes).padStart(2, '0'));
        mes++; if (mes > 12) { mes = 1; ano++; }
    }
    return meses;
}

function carregarDados() {
    if (!auth.currentUser) return;
    mostrarStatus('🔄 Carregando...', 'info');
    const finalizar = () => { try { popularFiltros(); aplicarFiltros(); } catch (e) { mostrarStatus('❌ Erro: ' + e.message, 'erro'); } };
    const timeout = setTimeout(() => {
        dados = DADOS_BASE.map(d => ({ ...d, mensal: MESES.reduce((acc, m) => { acc[m] = 0; return acc; }, {}) }));
        dadosCarregados = true; finalizar();
    }, 5000);

    database.ref(REF_CONTRATOS).once('value').then(snapshot => {
        clearTimeout(timeout);
        const val = snapshot.val();
        if (val && Object.keys(val).length > 0) {
            dados = Object.values(val).map(d => {
                if (!d.mensal) d.mensal = {};
                Object.keys(d.mensal).forEach(m => { if (!MESES.includes(m)) MESES.push(m); });
                MESES.forEach(m => { if (d.mensal[m] === undefined) d.mensal[m] = 0; });
                d.ultimaMedLancada = d.ultimaMedLancada || 0; d.mesUltimaMedLancada = d.mesUltimaMedLancada || '';
                d.ultimaMedPaga = d.ultimaMedPaga || 0; d.mesUltimaMedPaga = d.mesUltimaMedPaga || '';
                d.saldoEmpenho = d.saldoEmpenho || 0;
                return d;
            });
            MESES.sort(); dadosCarregados = true; finalizar();
        } else {
            database.ref(`${REF_CONFIG}/baseVazia`).once('value').then(s => {
                if (s.val() === true) { dados = []; dadosCarregados = true; finalizar(); }
                else {
                    dados = DADOS_BASE.map(d => ({ ...d, mensal: MESES.reduce((acc, m) => { acc[m] = 0; return acc; }, {}) }));
                    dadosCarregados = true; finalizar(); salvarDadosFirebase(true);
                }
            });
        }
    }).catch(() => {
        clearTimeout(timeout);
        dados = DADOS_BASE.map(d => ({ ...d, mensal: MESES.reduce((acc, m) => { acc[m] = 0; return acc; }, {}) }));
        dadosCarregados = true; finalizar();
    });
}

function salvarDadosFirebase(silencioso) {
    if (salvando) return;
    salvando = true;
    if (!silencioso) mostrarStatus('💾 Salvando...', 'info');
    const updates = {};
    dados.forEach(d => { updates[`${REF_CONTRATOS}/${d.id}`] = d; });
    updates[`${REF_CONFIG}/ultimaAtualizacao`] = new Date().toISOString();
    updates[`${REF_CONFIG}/usuario`] = auth.currentUser ? auth.currentUser.email : 'desconhecido';
    if (dados.length > 0) updates[`${REF_CONFIG}/baseVazia`] = null;

    database.ref().update(updates)
        .then(() => { salvando = false; mostrarStatus(silencioso ? '✅ Dados iniciais salvos!' : '✅ Dados salvos!'); })
        .catch(err => { salvando = false; mostrarStatus('❌ Erro: ' + err.message, 'erro'); });
}
function salvarTodos() { salvarDadosFirebase(false); }
function salvarItemFirebase(id) {
    clearTimeout(debounceTimers[id]);
    debounceTimers[id] = setTimeout(() => {
        const item = dados.find(d => d.id === id);
        if (item) database.ref(`${REF_CONTRATOS}/${id}`).set(item).then(() => mostrarStatus('💾 Salvo!')).catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
    }, 800);
}

function abrirModalUpload() {
    arquivoAtual = null; dadosPreview = null;
    document.getElementById('inputFile').value = '';
    document.getElementById('arquivoInfo').style.display = 'none';
    document.getElementById('previewUpload').style.display = 'none';
    document.getElementById('resultadoUpload').innerHTML = '';
    const btn = document.getElementById('btnProcessarUpload');
    btn.disabled = true; btn.style.background = '#94a3b8'; btn.textContent = '✅ Confirmar Importação';
    document.getElementById('modalUpload').classList.add('open');
}
function fecharModalUpload() { document.getElementById('modalUpload').classList.remove('open'); arquivoAtual = null; dadosPreview = null; }
function handleDrop(event) { event.preventDefault(); if (event.dataTransfer.files.length > 0) processarArquivo(event.dataTransfer.files[0]); }
function handleFileSelect(event) { if (event.target.files[0]) processarArquivo(event.target.files[0]); }

function processarArquivo(file) {
    if (!file.name.match(/\.xlsx?$/i)) { alert('⚠️ Selecione um arquivo .xlsx ou .xls'); return; }
    arquivoAtual = file;
    document.getElementById('arquivoInfo').style.display = 'block';
    document.getElementById('arquivoInfo').innerHTML = `⏳ Lendo arquivo...`;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellStyles: true });
            let aba = workbook.Sheets[XLSX_CONFIG.aba];
            if (!aba) {
                const abaSimilar = workbook.SheetNames.find(n => n.toLowerCase().includes('execu') || n.toLowerCase().includes('financeira'));
                if (abaSimilar) aba = workbook.Sheets[abaSimilar];
                else { alert(`❌ Aba "${XLSX_CONFIG.aba}" não encontrada.`); return; }
            }
            const matriz = XLSX.utils.sheet_to_json(aba, { header: 1, defval: '' });
            let mesesInfo = detectarMesesDoCabecalho(matriz);
            let mesesDetectados, colunasMeses, origemDeteccao;

            if (mesesInfo && mesesInfo.length > 0) {
                mesesDetectados = mesesInfo.map(m => m.mes); colunasMeses = mesesInfo.map(m => m.coluna);
                origemDeteccao = `Cabeçalho (linha ${XLSX_CONFIG.linhaCabecalho})`;
            } else {
                mesesDetectados = gerarMesesPadrao(); colunasMeses = mesesDetectados.map((_, i) => XLSX_CONFIG.colunaInicioFixa + i);
                origemDeteccao = 'Padrão (fallback)';
            }

            MESES = mesesDetectados; const dadosExtraidos = [];

            for (let i = XLSX_CONFIG.linhaInicio - 1; i <= XLSX_CONFIG.linhaFim - 1; i++) {
                const linha = matriz[i];
                if (!linha || (aba['!rows'] && aba['!rows'][i] && aba['!rows'][i].hidden)) continue;
                const excelRow = i + 1;
                const contratoBruto = String(linha[XLSX_CONFIG.colunaContrato] || '').trim();
                if (!contratoBruto) continue;
                const contrato = contratoBruto.replace(/\s+/g, '');

                let ulDefinida = 'A DEFINIR';
                if (excelRow >= 9 && excelRow <= 16) ulDefinida = 'SOBRAL';
                else if (excelRow >= 17 && excelRow <= 23) ulDefinida = 'ICÓ';
                else if (excelRow >= 24 && excelRow <= 34) ulDefinida = 'FORTALEZA';
                else if (excelRow >= 35 && excelRow <= 45) ulDefinida = 'RUSSAS';
                else if (excelRow >= 46 && excelRow <= 57) ulDefinida = 'BOA VIAGEM';
                else if (excelRow === 58) ulDefinida = 'SMT';

                const mensal = {};
                mesesDetectados.forEach((mes, idx) => {
                    const valor = linha[colunasMeses[idx]];
                    mensal[mes] = typeof valor === 'number' ? valor : parseNumero(valor) || 0;
                });

                dadosExtraidos.push({
                    linha: excelRow, contrato, ul: ulDefinida,
                    br: String(linha[XLSX_CONFIG.colunaBR] || '').trim(),
                    servico: String(linha[XLSX_CONFIG.colunaServico] || '').trim(),
                    contratada: String(linha[XLSX_CONFIG.colunaContratada] || '').trim(),
                    vigencia: parseDataExcel(linha[XLSX_CONFIG.colunaVigencia]),
                    ultimaMedLancada: parseInt(linha[XLSX_CONFIG.colunaUltimaMedLancada]) || 0,
                    mesUltimaMedLancada: parseMesAnoExcel(linha[XLSX_CONFIG.colunaMesUltimaMedLancada]),
                    ultimaMedPaga: parseInt(linha[XLSX_CONFIG.colunaUltimaMedPaga]) || 0,
                    mesUltimaMedPaga: parseMesAnoExcel(linha[XLSX_CONFIG.colunaMesUltimaMedPaga]),
                    saldoEmpenho: parseNumero(linha[XLSX_CONFIG.colunaSaldoEmpenho]), mensal
                });
            }

            dadosPreview = { contratos: dadosExtraidos, meses: mesesDetectados, origem: origemDeteccao };
            renderizarPreview(dadosPreview);
            const btn = document.getElementById('btnProcessarUpload');
            btn.disabled = false; btn.style.background = '#2c7a30';
            btn.textContent = `✅ Confirmar Importação (${dadosExtraidos.length} contratos)`;
        } catch (err) { alert('❌ Erro ao ler o arquivo: ' + err.message); }
    };
    reader.readAsArrayBuffer(file);
}

function renderizarPreview(dadosPreviewObj) {
    const wrap = document.getElementById('previewUpload');
    const head = document.getElementById('previewHead');
    const body = document.getElementById('previewBody');
    const { contratos, meses, origem } = dadosPreviewObj;

    document.getElementById('arquivoInfo').innerHTML = `📄 <strong>${arquivoAtual.name}</strong> (${(arquivoAtual.size / 1024).toFixed(1)} KB)<br>🔍 <strong>Detecção:</strong> ${origem}<br>📅 <strong>Período:</strong> ${meses[0]} a ${meses[meses.length - 1]} (${meses.length} meses)<br>📋 <strong>Contratos:</strong> ${contratos.length}`;

    if (contratos.length === 0) {
        wrap.style.display = 'block'; head.innerHTML = '';
        body.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:12px;color:#dc3545;">⚠️ Nenhum contrato encontrado.</td></tr>';
        return;
    }

    let headHtml = '<tr><th style="padding:6px;">Contrato</th><th style="padding:6px;">UL</th><th style="padding:6px;">Vigência</th><th style="padding:6px;">Últ.Lanç</th><th style="padding:6px;">Mês Lanç</th><th style="padding:6px;">Últ.Paga</th><th style="padding:6px;">Mês Paga</th><th style="padding:6px; color:#81c784;">Saldo Empenho</th>';
    const prim = meses.slice(0, 2), ult = meses.slice(-2);
    prim.forEach(m => { const [a, m2] = m.split('-'); headHtml += `<th style="padding:6px;">${m2}/${a.slice(2)}</th>`; });
    if (meses.length > 4) headHtml += '<th style="padding:6px;">...</th>';
    ult.forEach(m => { const [a, m2] = m.split('-'); headHtml += `<th style="padding:6px;">${m2}/${a.slice(2)}</th>`; });
    headHtml += '<th style="padding:6px;">Total Mensal</th></tr>';
    head.innerHTML = headHtml;

    body.innerHTML = '';
    contratos.slice(0, 5).forEach(d => {
        const total = meses.reduce((s, m) => s + (d.mensal[m] || 0), 0);
        let rowHtml = `<tr><td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:600;">${d.contrato}</td><td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#4496CB;">${d.ul}</td><td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;">${d.vigencia || '-'}</td><td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#2c7a30;">${d.ultimaMedLancada || '-'}</td><td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#2c7a30;">${d.mesUltimaMedLancada || '-'}</td><td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#2c7a30;">${d.ultimaMedPaga || '-'}</td><td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#2c7a30;">${d.mesUltimaMedPaga || '-'}</td><td style="padding:5px;text-align:right;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#1b5e20;">${formatarMoeda(d.saldoEmpenho)}</td>`;
        prim.forEach(m => { rowHtml += `<td style="padding:5px;text-align:right;border-bottom:1px solid #e2e8f0;">${formatarNumeroInput(d.mensal[m])}</td>`; });
        if (meses.length > 4) rowHtml += `<td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;color:#94a3b8;">…</td>`;
        ult.forEach(m => { rowHtml += `<td style="padding:5px;text-align:right;border-bottom:1px solid #e2e8f0;">${formatarNumeroInput(d.mensal[m])}</td>`; });
        rowHtml += `<td style="padding:5px;text-align:right;border-bottom:1px solid #e2e8f0;font-weight:600;color:#2c7a30;">${formatarMoeda(total)}</td></tr>`;
        body.innerHTML += rowHtml;
    });

    if (contratos.length > 5) body.innerHTML += `<tr><td colspan="${meses.length + 10}" style="text-align:center;padding:8px;color:#64748b;font-size:11px;">... e mais ${contratos.length - 5} contratos</td></tr>`;
    wrap.style.display = 'block';
}

function processarUpload() {
    if (!dadosPreview || dadosPreview.contratos.length === 0) { alert('⚠️ Nenhum dado para importar.'); return; }
    const resultadoEl = document.getElementById('resultadoUpload');
    resultadoEl.innerHTML = '<span style="color:#7c3aed;">⏳ Processando...</span>';

    const { contratos, meses: mesesDetectados } = dadosPreview;
    let atualizados = 0;
    const criadosBase = [], criadosSemDados = [], updates = {}, detalhes = [];

    contratos.forEach(item => {
        const nomeLimpo = limparNomeContrato(item.contrato);
        let contratoSis = dados.find(d => limparNomeContrato(d.contrato) === nomeLimpo);

        if (!contratoSis) {
            const molde = DADOS_BASE.find(d => limparNomeContrato(d.contrato) === nomeLimpo);
            if (molde) { contratoSis = { ...molde, mensal: {} }; criadosBase.push(item.contrato); }
            else {
                contratoSis = { id: 'c_' + Date.now() + '_' + criadosSemDados.length, ul: 'A DEFINIR', contrato: item.contrato, br: '', servico: '', contratada: '', situacao: 'Ativo', vigencia: '', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '', saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0, mensal: {} };
                criadosSemDados.push(item.contrato);
            }
            dados.push(contratoSis);
        }

        Object.assign(contratoSis, {
            ul: item.ul || contratoSis.ul || 'A DEFINIR', br: item.br || contratoSis.br || '', servico: item.servico || contratoSis.servico || '', contratada: item.contratada || contratoSis.contratada || '',
            vigencia: String(item.vigencia || '').trim(), ultimaMedLancada: parseInt(item.ultimaMedLancada) || 0, mesUltimaMedLancada: String(item.mesUltimaMedLancada || '').trim(),
            ultimaMedPaga: parseInt(item.ultimaMedPaga) || 0, mesUltimaMedPaga: String(item.mesUltimaMedPaga || '').trim(), saldoEmpenho: parseFloat(item.saldoEmpenho) || 0
        });

        if (!contratoSis.mensal) contratoSis.mensal = {};
        mesesDetectados.forEach(mes => { contratoSis.mensal[mes] = item.mensal[mes] || 0; });

        updates[`${REF_CONTRATOS}/${contratoSis.id}`] = contratoSis;
        atualizados++;
        detalhes.push(`✅ ${item.contrato} → UL:${contratoSis.ul} | Saldo Emp:${formatarMoeda(contratoSis.saldoEmpenho)}`);
    });

    updates[`${REF_CONFIG}/baseVazia`] = null;

    if (atualizados === 0) { resultadoEl.innerHTML = `❌ <strong>Nenhum contrato atualizado.</strong>`; return; }

    database.ref().update(updates).then(() => {
        resultadoEl.innerHTML = `<div style="background:#e8f5e9; padding:14px; border-radius:10px; border-left:4px solid #2c7a30; max-height:300px; overflow-y:auto;">✅ <strong>${atualizados} contratos importados!</strong><br>📅 Período: ${mesesDetectados[0]} a ${mesesDetectados[mesesDetectados.length - 1]}<details style="margin-top:10px;" open><summary style="cursor:pointer; font-weight:600;">Detalhes</summary><div style="margin-top:6px; font-size:11px; font-family:monospace; max-height:200px; overflow-y:auto;">${detalhes.join('<br>')}</div></details></div>`;
        popularFiltros(); aplicarFiltros(); mostrarStatus(`✅ ${atualizados} contratos importados!`);
    }).catch(err => { resultadoEl.innerHTML = `<span style="color:#dc3545;">❌ Erro: ${err.message}</span>`; });
}

function toggleTodosCheckbox() {
    const isChecked = document.getElementById('checkAll').checked;
    document.querySelectorAll('.checkbox-contrato').forEach(cb => { cb.checked = isChecked; cb.dispatchEvent(new Event('change')); });
}
function toggleSelecaoContrato(id, isChecked) {
    if (isChecked) selecionadosIds.add(id); else selecionadosIds.delete(id);
    atualizarBotaoExclusao();
}
function atualizarBotaoExclusao() {
    const btn = document.getElementById('btnExcluirSelecionados'), count = selecionadosIds.size;
    if (count > 0) { btn.disabled = false; btn.style.background = '#dc3545'; btn.innerHTML = `🗑️ Excluir Selecionados (${count})`; }
    else { btn.disabled = true; btn.style.background = '#94a3b8'; btn.innerHTML = `🗑️ Excluir Selecionados (0)`; }
}
function excluirSelecionados() {
    if (selecionadosIds.size === 0 || !confirm(`⚠️ Excluir os ${selecionadosIds.size} contratos selecionados?`)) return;
    mostrarStatus(`🗑️ Excluindo...`, 'info');
    const updates = {};
    selecionadosIds.forEach(id => { updates[`${REF_CONTRATOS}/${id}`] = null; dados = dados.filter(d => d.id !== id); });
    database.ref().update(updates).then(() => {
        selecionadosIds.clear(); document.getElementById('checkAll').checked = false; atualizarBotaoExclusao(); popularFiltros(); aplicarFiltros(); mostrarStatus('✅ Excluídos!');
    }).catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
}

function excluirContrato(id) {
    const item = dados.find(d => d.id === id); if (!item) return;
    document.getElementById('modalTitle').textContent = '⚠️ Excluir Contrato';
    document.getElementById('modalMessage').innerHTML = `Excluir <strong>${item.contrato}</strong> (${item.contratada})?`;
    const btn = document.getElementById('btnConfirmAction');
    btn.textContent = '🗑️ Excluir'; btn.onclick = () => confirmarExclusao(id);
    document.getElementById('modalConfirm').classList.add('open');
}
function confirmarExclusao(id) {
    fecharModal(); mostrarStatus('🗑️ Excluindo...', 'info');
    database.ref(`${REF_CONTRATOS}/${id}`).remove().then(() => {
        dados = dados.filter(d => d.id !== id); selecionadosIds.delete(id); atualizarBotaoExclusao(); popularFiltros(); aplicarFiltros(); mostrarStatus('✅ Excluído!');
    }).catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
}
function fecharModal() { document.getElementById('modalConfirm').classList.remove('open'); }

function abrirModalAdicionar() {
    ['addUL', 'addContrato', 'addBR', 'addServico', 'addContratada', 'addVigencia', 'addUltimaMedLancada', 'addMesUltimaMedLancada', 'addUltimaMedPaga', 'addMesUltimaMedPaga', 'addSaldoEmpenho', 'addRap', 'addLoa', 'addRemanejMenos', 'addRemanejMais'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('addSituacao').value = 'Ativo';
    document.getElementById('modalAdicionar').classList.add('open');
}
function fecharModalAdicionar() { document.getElementById('modalAdicionar').classList.remove('open'); }
function confirmarAdicionar() {
    const ul = document.getElementById('addUL').value.trim(), contrato = document.getElementById('addContrato').value.trim();
    if (!ul || !contrato) { alert('⚠️ Preencha UL e Contrato'); return; }

    const novo = {
        id: 'c_' + Date.now(), ul, contrato,
        br: document.getElementById('addBR').value.trim(), servico: document.getElementById('addServico').value.trim(), contratada: document.getElementById('addContratada').value.trim(),
        situacao: document.getElementById('addSituacao').value, vigencia: document.getElementById('addVigencia').value.trim(),
        ultimaMedLancada: parseInt(document.getElementById('addUltimaMedLancada').value) || 0, mesUltimaMedLancada: document.getElementById('addMesUltimaMedLancada').value.trim(),
        ultimaMedPaga: parseInt(document.getElementById('addUltimaMedPaga').value) || 0, mesUltimaMedPaga: document.getElementById('addMesUltimaMedPaga').value.trim(),
        saldoEmpenho: parseNumero(document.getElementById('addSaldoEmpenho').value), rap: parseNumero(document.getElementById('addRap').value), loa: parseNumero(document.getElementById('addLoa').value),
        remanejMenos: parseNumero(document.getElementById('addRemanejMenos').value), remanejMais: parseNumero(document.getElementById('addRemanejMais').value),
        mensal: MESES.reduce((acc, m) => { acc[m] = 0; return acc; }, {})
    };

    database.ref().update({ [`${REF_CONTRATOS}/${novo.id}`]: novo, [`${REF_CONFIG}/baseVazia`]: null })
        .then(() => { dados.push(novo); fecharModalAdicionar(); popularFiltros(); aplicarFiltros(); mostrarStatus('✅ Adicionado!'); })
        .catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
}

function ativarModoEdicao() {
    modoEdicao = !modoEdicao;
    const btn = document.getElementById('btnEditar');
    btn.textContent = modoEdicao ? '🔒 Desativar Edição' : '✏️ Modo Edição';
    btn.style.background = modoEdicao ? '#2c7a30' : '#e9a23b';
    btn.style.color = modoEdicao ? 'white' : '#1a1a1a';
    renderizarTabelaContratos(); renderizarTabelaMensal();
    mostrarStatus(modoEdicao ? '✏️ Edição ATIVADA' : '🔒 Edição desativada', 'info');
}

function popularFiltros() {
    const ctrs = [...new Set(dados.map(d => d.contratada).filter(Boolean))].sort();
    const brs = [...new Set(dados.map(d => d.br).filter(Boolean))].sort();
    const selC = document.getElementById('filtroContratada'), selB = document.getElementById('filtroBR');
    const ctrAt = selC.value, brAt = selB.value;
    selC.innerHTML = '<option value="todas">Todas as Empresas</option>' + ctrs.map(c => `<option value="${c}">${c}</option>`).join('');
    selB.innerHTML = '<option value="todas">Todas</option>' + brs.map(b => `<option value="${b}">${b}</option>`).join('');
    if (ctrs.includes(ctrAt)) selC.value = ctrAt;
    if (brs.includes(brAt)) selB.value = brAt;
}

function atualizarCardsUL() {
    const grid = document.getElementById('ulCardsGrid');
    const uls = [...new Set(dados.map(d => d.ul).filter(Boolean))].sort();
    const sit = document.getElementById('filtroSituacao').value, ctr = document.getElementById('filtroContratada').value, br = document.getElementById('filtroBR').value;
    grid.innerHTML = '';

    uls.forEach(u => {
        const cs = dados.filter(d => {
            if (d.ul !== u) return false;
            if (sit !== 'todas' && !(d.situacao || '').toLowerCase().includes(sit.toLowerCase())) return false;
            if (ctr !== 'todas' && d.contratada !== ctr) return false;
            if (br !== 'todas' && !(d.br || '').includes(br)) return false;
            return true;
        });
        if (cs.length === 0) return;

        const tMen = cs.reduce((a, d) => a + MESES.reduce((s, m) => s + ((d.mensal && d.mensal[m]) || 0), 0), 0);
        const tLoa = cs.reduce((a, d) => a + (d.loa || 0), 0);
        const tSal = cs.reduce((a, d) => a + (d.saldoEmpenho || 0), 0);
        const tRap = cs.reduce((a, d) => a + (d.rap || 0), 0);

        const card = document.createElement('div');
        card.className = `ul-card ${ulAtivaFiltro === u ? 'ativo' : ''}`;
        card.onclick = () => { ulAtivaFiltro = (ulAtivaFiltro === u) ? 'todas' : u; aplicarFiltros(); };
        card.innerHTML = `
            <div class="ul-card-title">${u} <span>${cs.length} contratos</span></div>
            <div class="ul-card-info">
                <div class="ul-info-row"><span>Total Mensal:</span><strong>${formatarMoedaCurta(tMen)}</strong></div>
                <div class="ul-info-row"><span>LOA:</span><strong>${formatarMoedaCurta(tLoa)}</strong></div>
                <div class="ul-info-row"><span>RAP:</span><strong>${formatarMoedaCurta(tRap)}</strong></div>
            </div>
            <div class="ul-info-saldo"><div class="label">Saldo Empenho</div><div class="value">${formatarMoedaCurta(tSal)}</div></div>`;
        grid.appendChild(card);
    });
}

function aplicarFiltros() {
    const sit = document.getElementById('filtroSituacao').value, ctr = document.getElementById('filtroContratada').value, br = document.getElementById('filtroBR').value;
    dadosFiltrados = dados.filter(item => {
        if (ulAtivaFiltro !== 'todas' && item.ul !== ulAtivaFiltro) return false;
        if (sit !== 'todas' && !(item.situacao || '').toLowerCase().includes(sit.toLowerCase())) return false;
        if (ctr !== 'todas' && item.contratada !== ctr) return false;
        if (br !== 'todas' && !(item.br || '').includes(br)) return false;
        return true;
    });
    atualizarCardsUL(); renderizarTabelaContratos(); renderizarTabelaMensal(); atualizarKPI();
}

function limparFiltros() {
    ulAtivaFiltro = 'todas';
    document.getElementById('filtroSituacao').value = 'todas';
    document.getElementById('filtroContratada').value = 'todas';
    document.getElementById('filtroBR').value = 'todas';
    aplicarFiltros();
}

function atualizarKPI() {
    const b = dadosFiltrados;
    const tLoa = b.reduce((a, d) => a + (d.loa || 0), 0), tRap = b.reduce((a, d) => a + (d.rap || 0), 0);
    const tSal = b.reduce((a, d) => a + (d.saldoEmpenho || 0), 0), tMen = b.reduce((a, d) => a + MESES.reduce((s, m) => s + ((d.mensal && d.mensal[m]) || 0), 0), 0);
    document.getElementById('totalSelecionado').innerHTML = `
        <span>📋 Contratos: <strong>${b.length}</strong></span> <span>💵 LOA: <strong>${formatarMoeda(tLoa)}</strong></span>
        <span>💰 RAP: <strong>${formatarMoeda(tRap)}</strong></span> <span>📅 Mensal: <strong>${formatarMoeda(tMen)}</strong></span>
        <span style="color:#4ade80;">🛡️ Saldo Empenho: <strong>${formatarMoeda(tSal)}</strong></span>`;
}

function renderizarTabelaContratos() {
    const tbody = document.getElementById('tabelaContratosBody');
    tbody.innerHTML = '';
    if (dadosFiltrados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="18" style="text-align:center;padding:20px;">Nenhum registro</td></tr>'; atualizarBotaoExclusao(); return;
    }

    dadosFiltrados.forEach(d => {
        const row = tbody.insertRow();
        const editClass = modoEdicao ? 'editavel' : '', readonly = modoEdicao ? '' : 'readonly';

        row.insertCell(0).innerHTML = `<input type="checkbox" class="checkbox-contrato" value="${d.id}" ${selecionadosIds.has(d.id) ? 'checked' : ''} onchange="toggleSelecaoContrato('${d.id}', this.checked)">`;
        row.insertCell(1).innerHTML = d.ul || '-'; row.insertCell(2).innerHTML = d.contrato || '-';
        row.insertCell(3).innerHTML = d.br || '-'; row.insertCell(4).innerHTML = d.servico || '-';
        row.insertCell(5).innerHTML = d.contratada || '-';

        const cSit = row.insertCell(6);
        if (modoEdicao) {
            cSit.innerHTML = `<select onchange="atualizarCampo('${d.id}', 'situacao', this.value)"><option value="Ativo" ${d.situacao === 'Ativo' ? 'selected' : ''}>Ativo</option><option value="Concluído" ${d.situacao === 'Concluído' ? 'selected' : ''}>Concluído</option><option value="Aguardando" ${(d.situacao || '').includes('Aguardando') ? 'selected' : ''}>Aguardando</option><option value="Cancelado" ${d.situacao === 'Cancelado' ? 'selected' : ''}>Cancelado</option><option value="" ${!d.situacao ? 'selected' : ''}>-</option></select>`;
        } else {
            cSit.innerHTML = `<span class="badge ${getBadgeClass(d.situacao)}">${d.situacao || 'N/I'}</span>`;
        }

        row.insertCell(7).innerHTML = `<input type="text" value="${d.vigencia || ''}" ${readonly} class="${editClass}" onchange="atualizarCampo('${d.id}','vigencia',this.value)" style="width:90px;" placeholder="DD/MM/AAAA">`;
        row.insertCell(8).innerHTML = `<input type="number" value="${d.ultimaMedLancada || 0}" ${readonly} class="${editClass}" onchange="atualizarCampoNumInt('${d.id}','ultimaMedLancada',this.value)" style="width:70px;">`;
        row.insertCell(9).innerHTML = `<input type="text" value="${d.mesUltimaMedLancada || ''}" ${readonly} class="${editClass}" onchange="atualizarCampo('${d.id}','mesUltimaMedLancada',this.value)" style="width:90px;" placeholder="jun/26">`;
        row.insertCell(10).innerHTML = `<input type="number" value="${d.ultimaMedPaga || 0}" ${readonly} class="${editClass}" onchange="atualizarCampoNumInt('${d.id}','ultimaMedPaga',this.value)" style="width:70px;">`;
        row.insertCell(11).innerHTML = `<input type="text" value="${d.mesUltimaMedPaga || ''}" ${readonly} class="${editClass}" onchange="atualizarCampo('${d.id}','mesUltimaMedPaga',this.value)" style="width:90px;" placeholder="mai/26">`;

        const cSal = row.insertCell(12);
        cSal.style.background = '#e8f5e9'; cSal.style.fontWeight = 'bold'; cSal.style.color = '#1b5e20';
        cSal.innerHTML = `<input type="text" value="${formatarNumeroInput(d.saldoEmpenho)}" ${readonly} class="${editClass}" onchange="atualizarCampoNum('${d.id}','saldoEmpenho',this.value)" style="width:110px; font-weight:bold; color:#1b5e20; border-color:#81c784;">`;

        row.insertCell(13).innerHTML = `<input type="text" value="${formatarNumeroInput(d.rap)}" ${readonly} class="${editClass}" onchange="atualizarCampoNum('${d.id}','rap',this.value)" style="width:100px;">`;
        row.insertCell(14).innerHTML = `<input type="text" value="${formatarNumeroInput(d.loa)}" ${readonly} class="${editClass}" onchange="atualizarCampoNum('${d.id}','loa',this.value)" style="width:110px;">`;
        row.insertCell(15).innerHTML = `<input type="text" value="${formatarNumeroInput(d.remanejMenos)}" ${readonly} class="${editClass}" onchange="atualizarCampoNum('${d.id}','remanejMenos',this.value)" style="width:100px;">`;
        row.insertCell(16).innerHTML = `<input type="text" value="${formatarNumeroInput(d.remanejMais)}" ${readonly} class="${editClass}" onchange="atualizarCampoNum('${d.id}','remanejMais',this.value)" style="width:100px;">`;

        row.insertCell(17).innerHTML = `<button class="acao-btn acao-btn-view" onclick="verDetalhes('${d.id}')" title="Ver">👁️</button><button class="acao-btn acao-btn-del" onclick="excluirContrato('${d.id}')" title="Excluir">🗑️</button>`;
    });

    atualizarBotaoExclusao();
    document.getElementById('checkAll').checked = dadosFiltrados.length > 0 && dadosFiltrados.every(d => selecionadosIds.has(d.id));
}

function renderizarTabelaMensal() {
    const thead = document.getElementById('tabelaMensalHead'), tbody = document.getElementById('tabelaMensalBody');
    let headHtml = '<th class="contrato-cell">Contrato</th><th>UL</th>';
    MESES.forEach(m => { const [a, m2] = m.split('-'); headHtml += `<th class="mes-col">${m2}/${a.slice(2)}</th>`; });
    headHtml += '<th>Total</th>';
    thead.innerHTML = headHtml;
    tbody.innerHTML = '';

    if (dadosFiltrados.length === 0) { tbody.innerHTML = `<tr><td colspan="${MESES.length + 3}" style="text-align:center;padding:20px;">Nenhum registro</td></tr>`; return; }

    dadosFiltrados.forEach(d => {
        const row = tbody.insertRow(), readonly = modoEdicao ? '' : 'readonly';
        const cC = row.insertCell(0);
        cC.className = 'contrato-cell';
        cC.innerHTML = `<strong>${d.contrato}</strong><br><small style="color:#64748b;">${(d.contratada || '').substring(0, 25)}${(d.contratada || '').length > 25 ? '...' : ''}</small>`;
        row.insertCell(1).innerHTML = d.ul || '-';

        let total = 0;
        MESES.forEach(m => {
            const v = (d.mensal && d.mensal[m]) || 0; total += v;
            row.insertCell().innerHTML = `<input type="text" value="${formatarNumeroInput(v)}" ${readonly} onchange="atualizarMensal('${d.id}','${m}',this.value)" placeholder="0,00">`;
        });
        const cT = row.insertCell(); cT.className = 'total-col'; cT.innerHTML = formatarMoeda(total);
    });
}

function atualizarCampo(id, campo, valor) { const item = dados.find(d => d.id === id); if (item) { item[campo] = valor; salvarItemFirebase(id); } }
function atualizarCampoNum(id, campo, valor) { const item = dados.find(d => d.id === id); if (item) { item[campo] = parseNumero(valor); salvarItemFirebase(id); aplicarFiltros(); } }
function atualizarCampoNumInt(id, campo, valor) { const item = dados.find(d => d.id === id); if (item) { item[campo] = parseInt(valor) || 0; salvarItemFirebase(id); } }

function atualizarMensal(id, mes, valor) {
    const item = dados.find(d => d.id === id); if (!item) return;
    if (!item.mensal) item.mensal = {};
    item.mensal[mes] = parseNumero(valor);
    const total = MESES.reduce((s, m) => s + ((item.mensal && item.mensal[m]) || 0), 0);
    document.querySelectorAll('#tabelaMensalBody tr').forEach(row => {
        if (row.cells[0] && row.cells[0].textContent.includes(item.contrato)) row.cells[row.cells.length - 1].innerHTML = formatarMoeda(total);
    });
    atualizarKPI(); atualizarCardsUL(); salvarItemFirebase(id);
}

function verDetalhes(id) {
    const item = dados.find(d => d.id === id); if (!item) return;
    const total = MESES.reduce((s, m) => s + ((item.mensal && item.mensal[m]) || 0), 0);
    alert(`📋 CONTRATO: ${item.contrato}\n🏢 UL: ${item.ul}\n🛣️ BR: ${item.br}\n📝 Serviço: ${item.servico}\n🏗️ Contratada: ${item.contratada}\n📌 Situação: ${item.situacao || 'N/I'}\n📅 Vigência: ${item.vigencia || '-'}\n🔢 Últ. Med. Lançada: ${item.ultimaMedLancada || '-'} (${item.mesUltimaMedLancada || '-'})\n🔢 Últ. Med. Paga: ${item.ultimaMedPaga || '-'} (${item.mesUltimaMedPaga || '-'})\n🛡️ Saldo Empenho: ${formatarMoeda(item.saldoEmpenho)}\n💰 RAP: ${formatarMoeda(item.rap)}\n💰 LOA: ${formatarMoeda(item.loa)}\n📊 Total Mensal: ${formatarMoeda(total)}`);
}

function resetarDados() {
    if (!confirm('⚠️ RESETAR tudo?')) return;
    mostrarStatus('🔄 Resetando...', 'info');
    database.ref().update({ [REF_CONTRATOS]: null, [REF_MEDICOES]: null, [`${REF_CONFIG}/baseVazia`]: null }).then(() => {
        dados = DADOS_BASE.map(d => ({ ...d, mensal: MESES.reduce((acc, m) => { acc[m] = 0; return acc; }, {}) }));
        dadosCarregados = true; salvarDadosFirebase(true); popularFiltros(); aplicarFiltros();
        medicoesSEI = []; popularFiltroContratoMedicoes(); aplicarFiltrosMedicoes();
        mostrarStatus('✅ Resetado!');
    }).catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
}

function esvaziarDados() {
    if (!confirm('⚠️ APAGAR TUDO?')) return;
    mostrarStatus('🧹 Apagando...', 'info');
    database.ref().update({ [REF_CONTRATOS]: null, [REF_MEDICOES]: null, [`${REF_CONFIG}/baseVazia`]: true, [`${REF_CONFIG}/ultimaAtualizacao`]: new Date().toISOString(), [`${REF_CONFIG}/usuario`]: auth.currentUser ? auth.currentUser.email : 'desconhecido' }).then(() => {
        dados = []; dadosCarregados = true; popularFiltros(); aplicarFiltros();
        medicoesSEI = []; popularFiltroContratoMedicoes(); aplicarFiltrosMedicoes();
        mostrarStatus('✅ Apagado!');
    }).catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
}

document.addEventListener('keydown', function(e) { if (e.key === 'Escape') { fecharModal(); fecharModalAdicionar(); fecharModalUpload(); } });
document.addEventListener('click', function(e) {
    if (e.target.classList.contains('modal-overlay')) {
        if (e.target.id === 'modalUpload') fecharModalUpload();
        else if (e.target.id === 'modalAdicionar') fecharModalAdicionar();
        else if (e.target.id === 'modalConfirm') fecharModal();
    }
});

function carregarMedicoesSEI() {
    database.ref(REF_MEDICOES).once('value').then(snapshot => {
        const val = snapshot.val();
        medicoesSEI = val ? Object.values(val) : [];
        popularFiltroContratoMedicoes(); aplicarFiltrosMedicoes(); mostrarStatus('📋 Medições carregadas');
    }).catch(() => mostrarStatus('❌ Erro ao carregar medições', 'erro'));
}

function salvarMedicoesSEI(silencioso) {
    if (!silencioso) mostrarStatus('💾 Salvando medições...', 'info');
    database.ref(REF_MEDICOES).set(medicoesSEI).then(() => mostrarStatus('✅ Medições salvas')).catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
}

function popularFiltroContratoMedicoes() {
    const sel = document.getElementById('filtroContratoMedicoes'); if (!sel) return;
    const ctrs = [...new Set(medicoesSEI.map(m => m.contrato).filter(Boolean))].sort();
    const atual = sel.value;
    sel.innerHTML = '<option value="todos">Todos os Contratos</option>' + ctrs.map(c => `<option value="${c}">${c}</option>`).join('');
    if (ctrs.includes(atual)) sel.value = atual;
}

function aplicarFiltrosMedicoes() {
    const ctr = (document.getElementById('filtroContratoMedicoes')?.value) || 'todos';
    const busca = (document.getElementById('buscaMedicoes')?.value || '').trim().toLowerCase();
    medicoesFiltradas = medicoesSEI.filter(m => {
        if (ctr !== 'todos' && m.contrato !== ctr) return false;
        if (busca) {
            const alvo = [m.contrato, m.numMedicao, m.procPagamento, m.procSuporte].join(' ').toLowerCase();
            if (!alvo.includes(busca)) return false;
        }
        return true;
    });
    renderizarMedicoesSEI();
}

function renderizarMedicoesSEI() {
    const tbody = document.getElementById('tabelaMedicoesBody'), stats = document.getElementById('medicoesStats');
    if (!tbody) return;

    if (medicoesFiltradas.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;padding:20px;color:#64748b;">Nenhuma medição encontrada.</td></tr>`;
        if (stats) stats.innerHTML = '📋 0 medições'; return;
    }

    const grupos = {};
    medicoesFiltradas.forEach(m => { if (!grupos[m.contrato]) grupos[m.contrato] = []; grupos[m.contrato].push(m); });

    let html = '';
    Object.keys(grupos).sort().forEach(contrato => {
        const itens = grupos[contrato].sort((a, b) => (parseInt(a.numMedicao) || 0) - (parseInt(b.numMedicao) || 0));
        const exp = contratosExpandidos.has(contrato);
        html += `<tr class="grupo-contrato" onclick="toggleGrupoContrato('${contrato}')">
            <td colspan="5" style="background:#1e293b;color:white;font-weight:600;cursor:pointer;padding:10px 12px;">
                <span style="margin-right:8px;">${exp ? '▼' : '▶'}</span> 📄 Contrato <strong>${contrato}</strong>
                <span style="float:right;font-weight:400;font-size:12px;opacity:0.8;">${itens.length} ${itens.length === 1 ? 'medição' : 'medições'}</span>
            </td></tr>`;
        if (exp) {
            itens.forEach(m => {
                const id = m._id || m.id || '';
                html += `<tr class="linha-medicao">
                    <td style="padding:6px 10px;text-align:center;font-weight:600;color:#4496CB;">${contrato}</td>
                    <td style="padding:6px 10px;text-align:center;font-weight:600;">${m.numMedicao || '-'}</td>
                    <td style="padding:6px 10px;"><div style="display:flex;align-items:center;gap:6px;justify-content:space-between;"><span style="font-family:monospace;font-size:12px;">${m.procPagamento || '-'}</span>${m.procPagamento ? `<button class="btn-copiar" onclick="event.stopPropagation();copiarProcesso('${m.procPagamento}')" title="Copiar processo">📋</button>` : ''}</div></td>
                    <td style="padding:6px 10px;"><div style="display:flex;align-items:center;gap:6px;justify-content:space-between;"><span style="font-family:monospace;font-size:12px;">${m.procSuporte || '-'}</span>${m.procSuporte ? `<button class="btn-copiar" onclick="event.stopPropagation();copiarProcesso('${m.procSuporte}')" title="Copiar processo">📋</button>` : ''}</div></td>
                    <td style="padding:6px 10px;text-align:center;"><button class="acao-btn acao-btn-view" onclick="event.stopPropagation();abrirModalEdicaoMedicao('${id}')" title="Editar">✏️</button></td>
                </tr>`;
            });
        }
    });
    tbody.innerHTML = html;
    if (stats) {
        const tot = new Set(medicoesFiltradas.map(m => m.contrato)).size;
        stats.innerHTML = `📋 <strong>${medicoesFiltradas.length}</strong> medições em <strong>${tot}</strong> contrato${tot > 1 ? 's' : ''}`;
    }
}

function copiarProcesso(numero) {
    if (!numero) return;
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(numero).then(() => mostrarStatus('📋 Copiado')).catch(() => fallbackCopiar(numero));
    } else fallbackCopiar(numero);
}
function fallbackCopiar(texto) {
    const ta = document.createElement('textarea');
    ta.value = texto; ta.style.position = 'fixed'; ta.style.left = '-9999px';
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); mostrarStatus('📋 Copiado'); } catch (e) { alert('Selecione manualmente: ' + texto); }
    document.body.removeChild(ta);
}

function toggleGrupoContrato(c) { if (contratosExpandidos.has(c)) contratosExpandidos.delete(c); else contratosExpandidos.add(c); renderizarMedicoesSEI(); }
function expandirTodosMedicoes() { medicoesFiltradas.forEach(m => contratosExpandidos.add(m.contrato)); renderizarMedicoesSEI(); }
function recolherTodosMedicoes() { contratosExpandidos.clear(); renderizarMedicoesSEI(); }

function excluirMedicao(id) {
    if (!id || !confirm('Excluir esta medição?')) return;
    const idx = medicoesSEI.findIndex(m => (m._id || m.id) === id);
    if (idx === -1) return;
    const rem = medicoesSEI.splice(idx, 1)[0];
    database.ref(`${REF_MEDICOES}/${id}`).remove().then(() => { popularFiltroContratoMedicoes(); aplicarFiltrosMedicoes(); mostrarStatus('✅ Excluída'); })
    .catch(err => { medicoesSEI.push(rem); mostrarStatus('❌ Erro: ' + err.message, 'erro'); });
}

function abrirModalMedicao() {
    idMedicaoEmEdicao = null;
    document.querySelector('#modalMedicao .modal-header h3').innerHTML = '➕ Adicionar Medição SEI';
    document.querySelector('#modalMedicao .modal-footer').innerHTML = `
        <button onclick="fecharModalMedicao()" style="background:#94a3b8;color:white;padding:10px 20px;border:none;border-radius:6px;cursor:pointer;">Cancelar</button>
        <button onclick="confirmarAdicionarMedicao()" style="background:#2c7a30;color:white;padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:600;">✅ Adicionar</button>`;
    const sel = document.getElementById('addMedContrato'); if (!sel) return;
    sel.innerHTML = '<option value="">Selecione...</option>' + [...new Set(dados.map(d => d.contrato).filter(Boolean))].sort().map(c => `<option value="${c}">${c}</option>`).join('');
    document.getElementById('addMedNumero').value = ''; document.getElementById('addMedProcPag').value = ''; document.getElementById('addMedProcSup').value = '';
    document.getElementById('modalMedicao').classList.add('open');
}
function fecharModalMedicao() { document.getElementById('modalMedicao').classList.remove('open'); }

function confirmarAdicionarMedicao() {
    const ctr = document.getElementById('addMedContrato').value.trim(), num = document.getElementById('addMedNumero').value.trim();
    const pPag = document.getElementById('addMedProcPag').value.trim(), pSup = document.getElementById('addMedProcSup').value.trim();
    if (!ctr || !num) { alert('⚠️ Preencha Contrato e Número'); return; }

    if (idMedicaoEmEdicao) {
        const med = medicoesSEI.find(m => (m._id || m.id) === idMedicaoEmEdicao);
        if (!med) return;
        Object.assign(med, { contrato: ctr, numMedicao: num, procPagamento: pPag, procSuporte: pSup });
        database.ref(`${REF_MEDICOES}/${idMedicaoEmEdicao}`).set(med).then(() => {
            idMedicaoEmEdicao = null; fecharModalMedicao(); popularFiltroContratoMedicoes(); aplicarFiltrosMedicoes(); mostrarStatus('✅ Atualizada');
        }).catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
        return;
    }

    const id = 'm_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
    const nova = { _id: id, contrato: ctr, numMedicao: num, procPagamento: pPag, procSuporte: pSup, dataCadastro: new Date().toISOString() };
    database.ref(`${REF_MEDICOES}/${id}`).set(nova).then(() => {
        medicoesSEI.push(nova); contratosExpandidos.add(ctr); fecharModalMedicao(); popularFiltroContratoMedicoes(); aplicarFiltrosMedicoes(); mostrarStatus('✅ Adicionada');
    }).catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
}

function abrirModalUploadMedicoes() {
    const input = document.getElementById('inputFileMedicoes'); if (input) input.value = '';
    const info = document.getElementById('infoUploadMedicoes'); if (info) { info.style.display = 'none'; info.innerHTML = ''; }
    const btn = document.getElementById('btnProcessarMedicoes'); if (btn) { btn.disabled = true; btn.style.background = '#94a3b8'; }
    document.getElementById('modalUploadMedicoes').classList.add('open');
}
function fecharModalUploadMedicoes() { document.getElementById('modalUploadMedicoes').classList.remove('open'); }
function handleFileMedicoes(event) { if (event.target.files[0]) processarArquivoMedicoes(event.target.files[0]); }

function processarArquivoMedicoes(file) {
    if (!file.name.match(/\.xlsx?$/i)) { alert('⚠️ Selecione um arquivo .xlsx ou .xls'); return; }
    const info = document.getElementById('infoUploadMedicoes'); info.style.display = 'block'; info.innerHTML = '⏳ Lendo...';

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellStyles: true });
            let aba = workbook.Sheets[XLSX_CONFIG_MEDICOES.aba];
            if (!aba) {
                const sim = workbook.SheetNames.find(n => n.toLowerCase().includes('medi') || n.toLowerCase().includes('sei'));
                if (sim) aba = workbook.Sheets[sim];
                else { alert(`❌ Aba não encontrada.`); return; }
            }
            const matriz = XLSX.utils.sheet_to_json(aba, { header: 1, defval: '' });
            const extraidas = []; let ctrAt = '';

            for (let i = 1; i < matriz.length; i++) {
                const linha = matriz[i]; if (!linha) continue;
                let ctr = String(linha[XLSX_CONFIG_MEDICOES.colContrato] || '').trim();
                const num = String(linha[XLSX_CONFIG_MEDICOES.colNumMedicao] || '').trim();
                const pPag = String(linha[XLSX_CONFIG_MEDICOES.colProcPagamento] || '').trim();
                const pSup = String(linha[XLSX_CONFIG_MEDICOES.colProcSuporte] || '').trim();

                if (ctr) ctrAt = ctr; else ctr = ctrAt;
                if (!ctr || !num || (!pPag && !pSup)) continue;

                extraidas.push({ _id: 'm_' + Date.now() + '_' + i + '_' + Math.random().toString(36).slice(2, 6), contrato: ctr.replace(/\s+/g, ''), numMedicao: num, procPagamento: pPag, procSuporte: pSup, dataCadastro: new Date().toISOString() });
            }

            if (extraidas.length === 0) { info.innerHTML = '⚠️ Nenhuma medição encontrada.'; return; }

            const ctrs = new Set(extraidas.map(m => m.contrato));
            info.innerHTML = `📄 <strong>${file.name}</strong><br>📋 <strong>${extraidas.length}</strong> medições em <strong>${ctrs.size}</strong> contratos`;
            const btn = document.getElementById('btnProcessarMedicoes');
            btn.disabled = false; btn.style.background = '#2c7a30'; btn.textContent = `✅ Confirmar (${extraidas.length})`;
            btn.onclick = () => confirmarImportacaoMedicoes(extraidas);
        } catch (err) { info.innerHTML = '❌ Erro: ' + err.message; }
    };
    reader.readAsArrayBuffer(file);
}

function confirmarImportacaoMedicoes(novas) {
    if (!confirm(`Importar ${novas.length} medições?\nAs existentes nestes contratos serão substituídas.`)) return;
    const ctrsNovos = new Set(novas.map(m => m.contrato));
    medicoesSEI = medicoesSEI.filter(m => !ctrsNovos.has(m.contrato)).concat(novas);

    const obj = {}; medicoesSEI.forEach(m => { obj[m._id] = m; });
    database.ref(REF_MEDICOES).set(obj).then(() => {
        fecharModalUploadMedicoes(); popularFiltroContratoMedicoes(); aplicarFiltrosMedicoes(); mostrarStatus(`✅ ${novas.length} importadas`);
    }).catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
}

function irParaMedicoes() {
    document.querySelectorAll('.secao').forEach(s => s.style.display = 'none');
    document.getElementById('secaoMedicoes').style.display = 'block';
    document.querySelector('.toolbar').style.setProperty('display', 'none', 'important');
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelector('[data-secao="medicoes"]')?.classList.add('active');
    if (medicoesSEI.length === 0) carregarMedicoesSEI(); else aplicarFiltrosMedicoes();
}

function irParaExecucao() {
    document.querySelectorAll('.secao').forEach(s => s.style.display = 'none');
    document.getElementById('secaoExecucao').style.display = 'block';
    document.querySelector('.toolbar').style.setProperty('display', 'flex', 'important');
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelector('[data-secao="execucao"]')?.classList.add('active');
}

const _carregarDados = carregarDados;
carregarDados = function() {
    _carregarDados();
    setTimeout(() => { if (auth.currentUser && medicoesSEI.length === 0) carregarMedicoesSEI(); }, 1500);
};

function abrirModalEdicaoMedicao(id) {
    const med = medicoesSEI.find(m => (m._id || m.id) === id); if (!med) return;
    idMedicaoEmEdicao = id;
    document.querySelector('#modalMedicao .modal-header h3').innerHTML = '✏️ Editar Medição SEI';

    const sel = document.getElementById('addMedContrato');
    const ctrs = [...new Set(dados.map(d => d.contrato).filter(Boolean))].sort();
    if (!ctrs.includes(med.contrato)) ctrs.push(med.contrato);
    sel.innerHTML = '<option value="">Selecione...</option>' + ctrs.sort().map(c => `<option value="${c}">${c}</option>`).join('');

    sel.value = med.contrato;
    document.getElementById('addMedNumero').value = med.numMedicao;
    document.getElementById('addMedProcPag').value = med.procPagamento || '';
    document.getElementById('addMedProcSup').value = med.procSuporte || '';

    document.querySelector('#modalMedicao .modal-footer').innerHTML = `
        <button onclick="excluirMedicaoDaEdicao()" style="background:#dc3545; color:white; padding:10px 20px; border:none; border-radius:6px; cursor:pointer; margin-right:auto;">🗑️ Excluir</button>
        <button onclick="fecharModalMedicao()" style="background:#94a3b8;color:white;padding:10px 20px;border:none;border-radius:6px;cursor:pointer;">Cancelar</button>
        <button onclick="confirmarAdicionarMedicao()" style="background:#2c7a30;color:white;padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:600;">✅ Salvar Alterações</button>`;

    document.getElementById('modalMedicao').classList.add('open');
}

function excluirMedicaoDaEdicao() {
    if (!idMedicaoEmEdicao || !confirm('Excluir permanentemente?')) return;
    const idx = medicoesSEI.findIndex(m => (m._id || m.id) === idMedicaoEmEdicao);
    if (idx === -1) return;
    const rem = medicoesSEI.splice(idx, 1)[0];
    database.ref(`${REF_MEDICOES}/${idMedicaoEmEdicao}`).remove().then(() => {
        idMedicaoEmEdicao = null; fecharModalMedicao(); popularFiltroContratoMedicoes(); aplicarFiltrosMedicoes(); mostrarStatus('✅ Excluída');
    }).catch(err => { medicoesSEI.push(rem); mostrarStatus('❌ Erro: ' + err.message, 'erro'); });
}

window.addEventListener('DOMContentLoaded', () => {
    irParaExecucao(); 
    const inputSenha = document.getElementById('password');
    if (inputSenha) inputSenha.setAttribute('type', 'password');
});
