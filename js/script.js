// ============================================================
// 🔥 FIREBASE
// ============================================================
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

// ============================================================
// 📊 MESES
// ============================================================
let MESES = [
    '2024-12', '2025-01', '2025-02', '2025-03', '2025-04', '2025-05',
    '2025-06', '2025-07', '2025-08', '2025-09', '2025-10', '2025-11',
    '2025-12', '2026-01', '2026-02', '2026-03', '2026-04', '2026-05',
    '2026-06', '2026-07', '2026-08', '2026-09'
];

// ============================================================
// 📁 XLSX CONFIG
// ============================================================
const XLSX_CONFIG = {
    aba: 'Execução Financeira',
    linhaInicio: 11,
    linhaFim: 58,
    colunaContrato: 1,
    colunaUL: 0,
    colunaBR: 2,
    colunaServico: 3,
    colunaContratada: 5,
    linhaCabecalho: 8,
    colunaInicioFixa: 28,
    colunasIgnorar: [29, 41, 42],
    maxColunasVazias: 3,
    colunaVigencia: 8,
    colunaUltimaMedLancada: 10,
    colunaMesUltimaMedLancada: 11,
    colunaUltimaMedPaga: 12,
    colunaMesUltimaMedPaga: 13,
    colunaSaldoEmpenho: 18
};

// ============================================================
// 📊 DADOS BASE
// ============================================================
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

// ============================================================
// 📊 ESTADO
// ============================================================
let dados = [];
let dadosFiltrados = [];
let modoEdicao = false;
let dadosCarregados = false;
let salvando = false;
let debounceTimers = {};
let arquivoAtual = null;
let dadosPreview = null;
let ulAtivaFiltro = 'todas';
let selecionadosIds = new Set();
let loginManualFeito = false;   // Controla se o login foi manual (para não deslogar no F5)
let loginEmProgresso = false;   // 🆕 Controla se um login está em andamento (evita o duplo clique)
let appAbertoPara = null;   // evita abrir o app duas vezes para o mesmo usuário

function limparNomeContrato(nome) {
    if (!nome) return '';
    return String(nome).toUpperCase().replace(/\s+/g, '').trim();
}

// ============================================================
// 🔐 LOGIN E CONTROLE DE ACESSO
// ============================================================
async function fazerLogin(event) {
    if (event) event.preventDefault();

    const btn = document.getElementById('btnLogin');
    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password').value;
    const errorEl = document.getElementById('loginError');

    if (btn.disabled) return false;

    errorEl.classList.remove('show');
    if (!username || !password) {
        errorEl.textContent = '❌ Preencha usuário e senha!';
        errorEl.classList.add('show');
        return false;
    }

    btn.disabled = true;
    btn.textContent = '⏳ Entrando...';
    loginEmProgresso = true;

    try {
        await auth.signOut();
        loginManualFeito = true;

        const email = username.includes('@') ? username : username + '@interno.local';
        const cred = await auth.signInWithEmailAndPassword(email, password);

        // ⭐ É ISTO que elimina o duplo clique: o próprio login abre o app
        await abrirAppParaUsuario(cred.user);
        return false;

    } catch (erro) {
        loginManualFeito = false;
        console.error('Erro login:', erro);
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

function fazerLogout() {
    if (!confirm('Deseja realmente sair?')) return;
    loginManualFeito = false; // Reseta a flag
    loginEmProgresso = false;
    appAbertoPara = null; 
    auth.signOut().then(() => {
        document.getElementById('appContainer').style.display = 'none';
        document.getElementById('modalTrocaSenha').style.display = 'none';
        document.getElementById('loginOverlay').style.display = 'flex';
        document.getElementById('username').value = '';
        document.getElementById('password').value = '';
        dadosCarregados = false;
    });
}

// Retrocompatível: aceita ser chamada como toggleSenha() ou toggleSenha(event)
function toggleSenha(btn) {
    const input = document.getElementById('password');
    const mostrar = input.type === 'password';
    input.type = mostrar ? 'text' : 'password';
    if (btn) btn.textContent = mostrar ? '🙈' : '👁️';
}


async function salvarNovaSenha(event) {
    if (event) event.preventDefault(); // Evita recarregar a página

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
        await database.ref(`${REF_USUARIOS}/${username}`).update({
            precisaTrocarSenha: false,
            dataTrocaSenha: new Date().toISOString()
        });

        alert('✅ Senha atualizada com sucesso! Você será redirecionado.');

        document.getElementById('modalTrocaSenha').style.display = 'none';
        document.getElementById('appContainer').style.display = 'block';
        if (!dadosCarregados) {
            carregarDados();
        }

    } catch (error) {
        console.error('Erro ao atualizar senha:', error);
        erroEl.textContent = '❌ Erro ao atualizar senha: ' + error.message;
        erroEl.style.display = 'block';
    }
}

auth.onAuthStateChanged(async (user) => {
    if (loginEmProgresso) return;            // o fazerLogin cuida do destino

    if (user && !loginManualFeito) {         // F5: exige login de novo
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
        dadosCarregados = false;
    }
});


// ============================================================
// 🔧 AUXILIARES
// ============================================================
function formatarMoeda(v) {
    if (!v && v !== 0) return 'R$ 0,00';
    return Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}
function formatarMoedaCurta(v) {
    if (!v && v !== 0) return 'R$ 0';
    if (v >= 1000000) return 'R$ ' + (v / 1000000).toFixed(2) + 'M';
    if (v >= 1000) return 'R$ ' + (v / 1000).toFixed(0) + 'K';
    return 'R$ ' + Number(v).toFixed(0);
}
function formatarNumeroInput(v) {
    if (!v && v !== 0) return '0,00';
    return Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function parseNumero(str) {
    if (typeof str === 'number') return str;
    if (!str) return 0;
    let s = String(str).trim();
    if (s.includes(',') && s.includes('.')) {
        s = s.replace(/\./g, '').replace(',', '.');
    } else if (s.includes(',')) {
        s = s.replace(',', '.');
    }
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
    el.textContent = msg;
    el.className = 'save-status';
    if (tipo === 'erro') el.classList.add('erro');
    else if (tipo === 'info') el.classList.add('info');
    el.style.display = 'block';
    clearTimeout(el._timeout);
    el._timeout = setTimeout(() => { el.style.display = 'none'; }, 3000);
}

// ============================================================
// 📅 PARSE DATAS
// ============================================================
function formatarDataBR(date) {
    return String(date.getDate()).padStart(2, '0') + '/' +
           String(date.getMonth() + 1).padStart(2, '0') + '/' +
           date.getFullYear();
}

function formatarMesAno(date) {
    const meses = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
    const mes = meses[date.getMonth()];
    const ano = String(date.getFullYear()).slice(2);
    return `${mes}/${ano}`;
}

function parseDataExcel(valor) {
    if (!valor && valor !== 0) return '';
    if (valor instanceof Date && !isNaN(valor)) return formatarDataBR(valor);
    if (typeof valor === 'number' && valor > 1 && valor < 100000) {
        try {
            const data = XLSX.SSF.parse_date_code(valor);
            if (data && data.y && data.m && data.d) {
                return String(data.d).padStart(2, '0') + '/' + String(data.m).padStart(2, '0') + '/' + data.y;
            }
        } catch (e) { }
    }
    const s = String(valor).trim();
    if (!s) return '';
    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return String(m[1]).padStart(2, '0') + '/' + String(m[2]).padStart(2, '0') + '/' + m[3];
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return String(m[3]).padStart(2, '0') + '/' + String(m[2]).padStart(2, '0') + '/' + m[1];
    m = s.match(/^([a-zç]{3,})[\-\/](\d{2,4})$/i);
    if (m) {
        const mapMeses = { jan: '01', fev: '02', mar: '03', abr: '04', mai: '05', jun: '06', jul: '07', ago: '08', set: '09', out: '10', nov: '11', dez: '12', janeiro: '01', fevereiro: '02', março: '03', abril: '04', maio: '05', junho: '06', julho: '07', agosto: '08', setembro: '09', outubro: '10', novembro: '11', dezembro: '12' };
        const nomeMes = m[1].toLowerCase();
        const mes = mapMeses[nomeMes];
        if (mes) {
            const ano = m[2].length === 2 ? '20' + m[2] : m[2];
            return mes + '/' + ano;
        }
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
                return meses[data.m - 1] + '/' + String(data.y).slice(2);
            }
        } catch (e) { }
    }
    const s = String(valor).trim();
    if (!s) return '';
    let m = s.match(/^([a-zç]{3,})[\-\/](\d{2,4})$/i);
    if (m) {
        const mapMeses = { jan: '01', fev: '02', mar: '03', abr: '04', mai: '05', jun: '06', jul: '07', ago: '08', set: '09', out: '10', nov: '11', dez: '12' };
        const nomeMes = m[1].toLowerCase().substring(0, 3);
        const mesNum = mapMeses[nomeMes];
        if (mesNum) {
            const ano = m[2].length === 2 ? m[2] : m[2].slice(-2);
            return nomeMes + '/' + ano;
        }
    }
    m = s.match(/^\d{1,2}\/(\d{1,2})\/(\d{4})/);
    if (m) {
        const mesNum = parseInt(m[1]);
        const ano = m[2].slice(-2);
        const meses = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
        return meses[mesNum - 1] + '/' + ano;
    }
    return s;
}

// ============================================================
// 📅 DETECÇÃO DE MESES
// ============================================================
function parseMesDoCabecalho(valor) {
    if (!valor) return null;
    if (valor instanceof Date && !isNaN(valor)) return valor.getFullYear() + '-' + String(valor.getMonth() + 1).padStart(2, '0');
    const s = String(valor).trim();
    if (!s) return null;
    let m = s.match(/^(\d{4})-(\d{1,2})(?:-\d{1,2})?/);
    if (m) {
        const ano = m[1]; const mes = String(parseInt(m[2])).padStart(2, '0');
        if (parseInt(mes) >= 1 && parseInt(mes) <= 12) return ano + '-' + mes;
    }
    m = s.match(/^\d{1,2}\/(\d{1,2})\/(\d{4})/);
    if (m) {
        const mes = String(parseInt(m[1])).padStart(2, '0'); const ano = m[2];
        if (parseInt(mes) >= 1 && parseInt(mes) <= 12) return ano + '-' + mes;
    }
    m = s.match(/([a-zç]{3,})[\/\s\-]+(\d{2,4})/i);
    if (m) {
        const mapMeses = { jan: '01', fev: '02', mar: '03', abr: '04', mai: '05', jun: '06', jul: '07', ago: '08', set: '09', out: '10', nov: '11', dez: '12', janeiro: '01', fevereiro: '02', março: '03', abril: '04', maio: '05', junho: '06', julho: '07', agosto: '08', setembro: '09', outubro: '10', novembro: '11', dezembro: '12' };
        const mes = mapMeses[m[1].toLowerCase()]; const ano = m[2].length === 2 ? '20' + m[2] : m[2];
        if (mes) return ano + '-' + mes;
    }
    return null;
}

function detectarMesesDoCabecalho(matriz) {
    const linhaCab = matriz[XLSX_CONFIG.linhaCabecalho - 1];
    if (!linhaCab) return null;
    const mesesDetectados = [];
    let colunasVaziasConsecutivas = 0;
    for (let j = XLSX_CONFIG.colunaInicioFixa; j < linhaCab.length; j++) {
        if (XLSX_CONFIG.colunasIgnorar.includes(j)) continue;
        const valor = linhaCab[j];
        const mes = parseMesDoCabecalho(valor);
        if (mes) {
            mesesDetectados.push({ coluna: j, mes: mes });
            colunasVaziasConsecutivas = 0;
        } else {
            colunasVaziasConsecutivas++;
            if (colunasVaziasConsecutivas >= XLSX_CONFIG.maxColunasVazias) break;
        }
    }
    return mesesDetectados.length > 0 ? mesesDetectados : null;
}

function gerarMesesPadrao() {
    const meses = []; let ano = 2024; let mes = 12;
    for (let i = 0; i < 22; i++) {
        meses.push(ano + '-' + String(mes).padStart(2, '0'));
        mes++;
        if (mes > 12) { mes = 1; ano++; }
    }
    return meses;
}

// ============================================================
// 🔥 CARREGAR DADOS
// ============================================================
function carregarDados() {
    if (!auth.currentUser) return;
    mostrarStatus('🔄 Carregando do Firebase...', 'info');

    function finalizarCarregamento() {
        try {
            popularFiltros();
            aplicarFiltros();
        } catch (e) {
            console.error('❌ Erro na renderização:', e);
            mostrarStatus('❌ Erro ao renderizar: ' + e.message, 'erro');
        }
    }

    const timeoutId = setTimeout(() => {
        dados = DADOS_BASE.map(d => ({ ...d, mensal: MESES.reduce((acc, m) => { acc[m] = 0; return acc; }, {}) }));
        dadosCarregados = true;
        finalizarCarregamento();
    }, 5000);

    database.ref(REF_CONTRATOS).once('value')
        .then(snapshot => {
            clearTimeout(timeoutId);
            const val = snapshot.val();
            if (val && Object.keys(val).length > 0) {
                dados = Object.values(val).map(d => {
                    if (!d.mensal) d.mensal = {};
                    Object.keys(d.mensal).forEach(m => { if (!MESES.includes(m)) MESES.push(m); });
                    MESES.forEach(m => { if (d.mensal[m] === undefined) d.mensal[m] = 0; });
                    if (d.ultimaMedLancada === undefined) d.ultimaMedLancada = 0;
                    if (d.mesUltimaMedLancada === undefined) d.mesUltimaMedLancada = '';
                    if (d.ultimaMedPaga === undefined) d.ultimaMedPaga = 0;
                    if (d.mesUltimaMedPaga === undefined) d.mesUltimaMedPaga = '';
                    if (d.saldoEmpenho === undefined) d.saldoEmpenho = 0;
                    return d;
                });
                MESES.sort();
                dadosCarregados = true;
                finalizarCarregamento();
            } else {
                database.ref(`${REF_CONFIG}/baseVazia`).once('value').then(s => {
                    if (s.val() === true) {
                        dados = [];
                        dadosCarregados = true;
                        finalizarCarregamento();
                    } else {
                        dados = DADOS_BASE.map(d => ({ ...d, mensal: MESES.reduce((acc, m) => { acc[m] = 0; return acc; }, {}) }));
                        dadosCarregados = true;
                        finalizarCarregamento();
                        salvarDadosFirebase(true);
                    }
                });
            }
        })
        .catch(err => {
            clearTimeout(timeoutId);
            dados = DADOS_BASE.map(d => ({ ...d, mensal: MESES.reduce((acc, m) => { acc[m] = 0; return acc; }, {}) }));
            dadosCarregados = true;
            finalizarCarregamento();
        });
}

// ============================================================
// 🔥 SALVAR
// ============================================================
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
        .catch(err => { salvando = false; mostrarStatus('❌ Erro ao salvar: ' + err.message, 'erro'); });
}
function salvarTodos() { salvarDadosFirebase(false); }
function salvarItemFirebase(id) {
    clearTimeout(debounceTimers[id]);
    debounceTimers[id] = setTimeout(() => {
        const item = dados.find(d => d.id === id);
        if (!item) return;
        database.ref(`${REF_CONTRATOS}/${id}`).set(item)
            .then(() => mostrarStatus('💾 Salvo!'))
            .catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
    }, 800);
}

// ============================================================
// 📁 UPLOAD XLSX
// ============================================================
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
function handleDrop(event) { event.preventDefault(); const files = event.dataTransfer.files; if (files.length > 0) processarArquivo(files[0]); }
function handleFileSelect(event) { const file = event.target.files[0]; if (file) processarArquivo(file); }

function processarArquivo(file) {
    if (!file.name.match(/\.xlsx?$/i)) { alert('⚠️ Selecione um arquivo .xlsx ou .xls'); return; }
    arquivoAtual = file;
    document.getElementById('arquivoInfo').style.display = 'block';
    document.getElementById('arquivoInfo').innerHTML = `⏳ Lendo arquivo...`;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            let aba = workbook.Sheets[XLSX_CONFIG.aba];
            if (!aba) {
                const abasDisponiveis = workbook.SheetNames;
                const abaSimilar = abasDisponiveis.find(n => n.toLowerCase().includes('execu') || n.toLowerCase().includes('financeira'));
                if (abaSimilar) { aba = workbook.Sheets[abaSimilar]; }
                else { alert(`❌ Aba "${XLSX_CONFIG.aba}" não encontrada.\n\nAbas: ${abasDisponiveis.join(', ')}`); return; }
            }
            const matriz = XLSX.utils.sheet_to_json(aba, { header: 1, defval: '' });
            let mesesInfo = detectarMesesDoCabecalho(matriz);
            let mesesDetectados, colunasMeses, origemDeteccao;

            if (mesesInfo && mesesInfo.length > 0) {
                mesesDetectados = mesesInfo.map(m => m.mes);
                colunasMeses = mesesInfo.map(m => m.coluna);
                origemDeteccao = `Cabeçalho (linha ${XLSX_CONFIG.linhaCabecalho})`;
            } else {
                mesesDetectados = gerarMesesPadrao();
                colunasMeses = mesesDetectados.map((_, i) => XLSX_CONFIG.colunaInicioFixa + i);
                origemDeteccao = 'Padrão (fallback)';
            }

            MESES = mesesDetectados;
            const dadosExtraidos = [];

            for (let i = XLSX_CONFIG.linhaInicio - 1; i <= XLSX_CONFIG.linhaFim - 1; i++) {
                const linha = matriz[i];
                if (!linha) continue;
                const excelRow = i + 1;
                if (aba['!rows'] && aba['!rows'][i] && aba['!rows'][i].hidden) { continue; }

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

                const vigencia = parseDataExcel(linha[XLSX_CONFIG.colunaVigencia]);
                const brLido = String(linha[XLSX_CONFIG.colunaBR] || '').trim();
                const servicoLido = String(linha[XLSX_CONFIG.colunaServico] || '').trim();
                const contratadaLida = String(linha[XLSX_CONFIG.colunaContratada] || '').trim();
                const ultimaMedLancada = parseInt(linha[XLSX_CONFIG.colunaUltimaMedLancada]) || 0;
                const mesUltimaMedLancada = parseMesAnoExcel(linha[XLSX_CONFIG.colunaMesUltimaMedLancada]);
                const ultimaMedPaga = parseInt(linha[XLSX_CONFIG.colunaUltimaMedPaga]) || 0;
                const mesUltimaMedPaga = parseMesAnoExcel(linha[XLSX_CONFIG.colunaMesUltimaMedPaga]);
                let saldoEmpenhoValor = parseNumero(linha[XLSX_CONFIG.colunaSaldoEmpenho]);

                const mensal = {};
                mesesDetectados.forEach((mes, idx) => {
                    const colIdx = colunasMeses[idx];
                    let valor = linha[colIdx];
                    if (typeof valor === 'string') valor = parseNumero(valor);
                    else if (typeof valor !== 'number') valor = 0;
                    mensal[mes] = valor || 0;
                });

                dadosExtraidos.push({
                    linha: excelRow, contrato: contrato, ul: ulDefinida, br: brLido, servico: servicoLido, contratada: contratadaLida,
                    vigencia: vigencia, ultimaMedLancada: ultimaMedLancada, mesUltimaMedLancada: mesUltimaMedLancada,
                    ultimaMedPaga: ultimaMedPaga, mesUltimaMedPaga: mesUltimaMedPaga, saldoEmpenho: saldoEmpenhoValor, mensal: mensal
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
    const dadosExtraidos = dadosPreviewObj.contratos;
    const meses = dadosPreviewObj.meses;
    const origem = dadosPreviewObj.origem;

    document.getElementById('arquivoInfo').innerHTML = `
        📄 <strong>${arquivoAtual.name}</strong> (${(arquivoAtual.size / 1024).toFixed(1)} KB)<br>
        🔍 <strong>Detecção:</strong> ${origem}<br>
        📅 <strong>Período:</strong> ${meses[0]} a ${meses[meses.length - 1]} (${meses.length} meses)<br>
        📋 <strong>Contratos encontrados:</strong> ${dadosExtraidos.length}
    `;

    if (dadosExtraidos.length === 0) {
        wrap.style.display = 'block'; head.innerHTML = '';
        body.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:12px;color:#dc3545;">⚠️ Nenhum contrato encontrado.</td></tr>';
        return;
    }

    let headHtml = '<tr><th style="padding:6px;">Contrato</th><th style="padding:6px;">UL</th><th style="padding:6px;">Vigência</th><th style="padding:6px;">Últ.Lanç</th><th style="padding:6px;">Mês Lanç</th><th style="padding:6px;">Últ.Paga</th><th style="padding:6px;">Mês Paga</th><th style="padding:6px; color:#81c784;">Saldo Empenho</th>';
    const primeirosMeses = meses.slice(0, 2);
    const ultimosMeses = meses.slice(-2);

    primeirosMeses.forEach(m => { const [ano, mes] = m.split('-'); headHtml += `<th style="padding:6px;">${mes}/${ano.slice(2)}</th>`; });
    if (meses.length > 4) headHtml += '<th style="padding:6px;">...</th>';
    ultimosMeses.forEach(m => { const [ano, mes] = m.split('-'); headHtml += `<th style="padding:6px;">${mes}/${ano.slice(2)}</th>`; });
    headHtml += '<th style="padding:6px;">Total Mensal</th></tr>';
    head.innerHTML = headHtml;

    body.innerHTML = '';
    dadosExtraidos.slice(0, 5).forEach(d => {
        const total = meses.reduce((s, m) => s + (d.mensal[m] || 0), 0);
        let rowHtml = `<tr>`;
        rowHtml += `<td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:600;">${d.contrato}</td>`;
        rowHtml += `<td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#4496CB;">${d.ul}</td>`;
        rowHtml += `<td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;">${d.vigencia || '-'}</td>`;
        rowHtml += `<td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#2c7a30;">${d.ultimaMedLancada || '-'}</td>`;
        rowHtml += `<td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#2c7a30;">${d.mesUltimaMedLancada || '-'}</td>`;
        rowHtml += `<td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#2c7a30;">${d.ultimaMedPaga || '-'}</td>`;
        rowHtml += `<td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#2c7a30;">${d.mesUltimaMedPaga || '-'}</td>`;
        rowHtml += `<td style="padding:5px;text-align:right;border-bottom:1px solid #e2e8f0;font-weight:bold;color:#1b5e20;">${formatarMoeda(d.saldoEmpenho)}</td>`;

        primeirosMeses.forEach(m => { rowHtml += `<td style="padding:5px;text-align:right;border-bottom:1px solid #e2e8f0;">${formatarNumeroInput(d.mensal[m])}</td>`; });
        if (meses.length > 4) rowHtml += `<td style="padding:5px;text-align:center;border-bottom:1px solid #e2e8f0;color:#94a3b8;">…</td>`;
        ultimosMeses.forEach(m => { rowHtml += `<td style="padding:5px;text-align:right;border-bottom:1px solid #e2e8f0;">${formatarNumeroInput(d.mensal[m])}</td>`; });

        rowHtml += `<td style="padding:5px;text-align:right;border-bottom:1px solid #e2e8f0;font-weight:600;color:#2c7a30;">${formatarMoeda(total)}</td>`;
        rowHtml += '</tr>';
        body.innerHTML += rowHtml;
    });

    if (dadosExtraidos.length > 5) {
        body.innerHTML += `<tr><td colspan="${meses.length + 10}" style="text-align:center;padding:8px;color:#64748b;font-size:11px;">... e mais ${dadosExtraidos.length - 5} contratos</td></tr>`;
    }
    wrap.style.display = 'block';
}

function processarUpload() {
    if (!dadosPreview || !dadosPreview.contratos || dadosPreview.contratos.length === 0) { alert('⚠️ Nenhum dado para importar.'); return; }
    const resultadoEl = document.getElementById('resultadoUpload');
    resultadoEl.innerHTML = '<span style="color:#7c3aed;">⏳ Processando...</span>';

    const dadosExtraidos = dadosPreview.contratos;
    const mesesDetectados = dadosPreview.meses;
    let atualizados = 0;
    let naoEncontrados = [];
    const criadosDoBase = []; const criadosSemDados = []; const updates = {}; const detalhes = [];

    dadosExtraidos.forEach(item => {
        const nomeLimpoExcel = limparNomeContrato(item.contrato);
        let contratoSistema = dados.find(d => limparNomeContrato(d.contrato) === nomeLimpoExcel);

        if (!contratoSistema) {
            const molde = DADOS_BASE.find(d => limparNomeContrato(d.contrato) === nomeLimpoExcel);
            if (molde) {
                contratoSistema = { ...molde, mensal: {} };
                criadosDoBase.push(item.contrato);
            } else {
                contratoSistema = {
                    id: 'c_' + Date.now() + '_' + criadosSemDados.length, ul: 'A DEFINIR', contrato: item.contrato, br: '', servico: '', contratada: '',
                    situacao: 'Ativo', vigencia: '', ultimaMedLancada: 0, mesUltimaMedLancada: '', ultimaMedPaga: 0, mesUltimaMedPaga: '',
                    saldoEmpenho: 0, rap: 0, loa: 0, remanejMenos: 0, remanejMais: 0, mensal: {}
                };
                criadosSemDados.push(item.contrato);
            }
            dados.push(contratoSistema);
        }

        contratoSistema.ul = item.ul || contratoSistema.ul || 'A DEFINIR';
        contratoSistema.br = item.br || contratoSistema.br || '';
        contratoSistema.servico = item.servico || contratoSistema.servico || '';
        contratoSistema.contratada = item.contratada || contratoSistema.contratada || '';
        contratoSistema.vigencia = String(item.vigencia || '').trim();
        contratoSistema.ultimaMedLancada = parseInt(item.ultimaMedLancada) || 0;
        contratoSistema.mesUltimaMedLancada = String(item.mesUltimaMedLancada || '').trim();
        contratoSistema.ultimaMedPaga = parseInt(item.ultimaMedPaga) || 0;
        contratoSistema.mesUltimaMedPaga = String(item.mesUltimaMedPaga || '').trim();
        contratoSistema.saldoEmpenho = parseFloat(item.saldoEmpenho) || 0;

        if (!contratoSistema.mensal) contratoSistema.mensal = {};
        let totalContrato = 0;
        mesesDetectados.forEach(mes => {
            contratoSistema.mensal[mes] = item.mensal[mes] || 0;
            totalContrato += item.mensal[mes] || 0;
        });

        updates[`${REF_CONTRATOS}/${contratoSistema.id}`] = contratoSistema;
        atualizados++;
        detalhes.push(`✅ ${item.contrato} → UL:${contratoSistema.ul} | Saldo Emp:${formatarMoeda(contratoSistema.saldoEmpenho)} | Lanç:${contratoSistema.ultimaMedLancada}/${contratoSistema.mesUltimaMedLancada || '-'}`);
    });

    updates[`${REF_CONFIG}/baseVazia`] = null;

    if (atualizados === 0) {
        resultadoEl.innerHTML = `<div style="background:#fef2f2; padding:12px; border-radius:8px; border-left:4px solid #dc3545;">❌ <strong>Nenhum contrato foi atualizado.</strong></div>`;
        return;
    }

    database.ref().update(updates)
        .then(() => {
            resultadoEl.innerHTML = `
                <div style="background:#e8f5e9; padding:14px; border-radius:10px; border-left:4px solid #2c7a30; max-height:300px; overflow-y:auto;">
                    ✅ <strong>${atualizados} contratos importados!</strong>
                    ${(criadosDoBase.length + criadosSemDados.length) ? `<br>➕ ${criadosDoBase.length + criadosSemDados.length} criados` : ''}
                    ${criadosSemDados.length ? `<br>⚠️ <strong>${criadosSemDados.length} sem cadastro base</strong>` : ''}<br>
                    📅 Período: <strong>${mesesDetectados[0]}</strong> a <strong>${mesesDetectados[mesesDetectados.length - 1]}</strong>
                    <details style="margin-top:10px;" open>
                        <summary style="cursor:pointer; font-weight:600;">Detalhes (${atualizados})</summary>
                        <div style="margin-top:6px; font-size:11px; font-family:monospace; max-height:200px; overflow-y:auto;">${detalhes.join('<br>')}</div>
                    </details>
                </div>`;
            popularFiltros();
            aplicarFiltros();
            mostrarStatus(`✅ ${atualizados} contratos importados!`);
        })
        .catch(err => { resultadoEl.innerHTML = `<span style="color:#dc3545;">❌ Erro: ${err.message}</span>`; });
}

// ============================================================
// MULTISELECT E EXCLUSÃO
// ============================================================
function toggleTodosCheckbox() {
    const isChecked = document.getElementById('checkAll').checked;
    const checkboxes = document.querySelectorAll('.checkbox-contrato');
    checkboxes.forEach(cb => { cb.checked = isChecked; cb.dispatchEvent(new Event('change')); });
}
function toggleSelecaoContrato(id, isChecked) {
    if (isChecked) { selecionadosIds.add(id); } else { selecionadosIds.delete(id); }
    atualizarBotaoExclusao();
}
function atualizarBotaoExclusao() {
    const btn = document.getElementById('btnExcluirSelecionados');
    const count = selecionadosIds.size;
    if (count > 0) {
        btn.disabled = false; btn.style.background = '#dc3545'; btn.style.color = 'white';
        btn.innerHTML = `🗑️ Excluir Selecionados (${count})`;
    } else {
        btn.disabled = true; btn.style.background = '#94a3b8'; btn.style.color = 'white';
        btn.innerHTML = `🗑️ Excluir Selecionados (0)`;
    }
}
function excluirSelecionados() {
    if (selecionadosIds.size === 0) return;
    if (!confirm(`⚠️ Excluir os ${selecionadosIds.size} contratos selecionados?`)) return;
    mostrarStatus(`🗑️ Excluindo...`, 'info');
    const updates = {};
    selecionadosIds.forEach(id => { updates[`${REF_CONTRATOS}/${id}`] = null; dados = dados.filter(d => d.id !== id); });
    database.ref().update(updates)
        .then(() => {
            selecionadosIds.clear();
            document.getElementById('checkAll').checked = false;
            atualizarBotaoExclusao();
            popularFiltros();
            aplicarFiltros();
            mostrarStatus('✅ Contratos excluídos!');
        })
        .catch(err => { mostrarStatus('❌ Erro: ' + err.message, 'erro'); });
}

// ============================================================
// 🗑️ EXCLUIR (Individual)
// ============================================================
function excluirContrato(id) {
    const item = dados.find(d => d.id === id);
    if (!item) return;
    document.getElementById('modalTitle').textContent = '⚠️ Excluir Contrato';
    document.getElementById('modalMessage').innerHTML = `Tem certeza que deseja excluir <strong>${item.contrato}</strong> (${item.contratada})?`;
    const btn = document.getElementById('btnConfirmAction');
    btn.textContent = '🗑️ Excluir';
    btn.onclick = () => confirmarExclusao(id);
    document.getElementById('modalConfirm').classList.add('open');
}
function confirmarExclusao(id) {
    fecharModal(); mostrarStatus('🗑️ Excluindo...', 'info');
    database.ref(`${REF_CONTRATOS}/${id}`).remove()
        .then(() => {
            dados = dados.filter(d => d.id !== id);
            selecionadosIds.delete(id);
            atualizarBotaoExclusao();
            popularFiltros(); aplicarFiltros(); mostrarStatus('✅ Excluído!');
        })
        .catch(err => { mostrarStatus('❌ Erro: ' + err.message, 'erro'); });
}
function fecharModal() { document.getElementById('modalConfirm').classList.remove('open'); }

// ============================================================
// ➕ ADICIONAR
// ============================================================
function abrirModalAdicionar() {
    ['addUL', 'addContrato', 'addBR', 'addServico', 'addContratada', 'addVigencia', 'addUltimaMedLancada', 'addMesUltimaMedLancada', 'addUltimaMedPaga', 'addMesUltimaMedPaga', 'addSaldoEmpenho', 'addRap', 'addLoa', 'addRemanejMenos', 'addRemanejMais'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('addSituacao').value = 'Ativo';
    document.getElementById('modalAdicionar').classList.add('open');
}
function fecharModalAdicionar() { document.getElementById('modalAdicionar').classList.remove('open'); }
function confirmarAdicionar() {
    const ul = document.getElementById('addUL').value.trim();
    const contrato = document.getElementById('addContrato').value.trim();
    if (!ul || !contrato) { alert('⚠️ Preencha UL e Contrato'); return; }

    const novo = {
        id: 'c_' + Date.now(), ul: ul, contrato: contrato,
        br: document.getElementById('addBR').value.trim(), servico: document.getElementById('addServico').value.trim(), contratada: document.getElementById('addContratada').value.trim(),
        situacao: document.getElementById('addSituacao').value, vigencia: document.getElementById('addVigencia').value.trim(),
        ultimaMedLancada: parseInt(document.getElementById('addUltimaMedLancada').value) || 0, mesUltimaMedLancada: document.getElementById('addMesUltimaMedLancada').value.trim(),
        ultimaMedPaga: parseInt(document.getElementById('addUltimaMedPaga').value) || 0, mesUltimaMedPaga: document.getElementById('addMesUltimaMedPaga').value.trim(),
        saldoEmpenho: parseNumero(document.getElementById('addSaldoEmpenho').value), rap: parseNumero(document.getElementById('addRap').value), loa: parseNumero(document.getElementById('addLoa').value),
        remanejMenos: parseNumero(document.getElementById('addRemanejMenos').value), remanejMais: parseNumero(document.getElementById('addRemanejMais').value),
        mensal: MESES.reduce((acc, m) => { acc[m] = 0; return acc; }, {})
    };

    const updatesAdd = {}; updatesAdd[`${REF_CONTRATOS}/${novo.id}`] = novo; updatesAdd[`${REF_CONFIG}/baseVazia`] = null;
    database.ref().update(updatesAdd)
        .then(() => { dados.push(novo); fecharModalAdicionar(); popularFiltros(); aplicarFiltros(); mostrarStatus('✅ Adicionado!'); })
        .catch(err => { mostrarStatus('❌ Erro: ' + err.message, 'erro'); });
}

// ============================================================
// ✏️ MODO EDIÇÃO
// ============================================================
function ativarModoEdicao() {
    modoEdicao = !modoEdicao;
    const btn = document.getElementById('btnEditar');
    btn.textContent = modoEdicao ? '🔒 Desativar Edição' : '✏️ Modo Edição';
    btn.style.background = modoEdicao ? '#2c7a30' : '#e9a23b';
    btn.style.color = modoEdicao ? 'white' : '#1a1a1a';
    renderizarTabelaContratos();
    renderizarTabelaMensal();
    mostrarStatus(modoEdicao ? '✏️ Modo edição ATIVADO' : '🔒 Modo edição desativado', 'info');
}

// ============================================================
// 🔽 FILTROS E CARDS DE UL
// ============================================================
function popularFiltros() {
    const contratadas = [...new Set(dados.map(d => d.contratada).filter(Boolean))].sort();
    const brs = [...new Set(dados.map(d => d.br).filter(Boolean))].sort();
    const selContratada = document.getElementById('filtroContratada');
    const selBR = document.getElementById('filtroBR');
    const ctrAtual = selContratada.value; const brAtual = selBR.value;

    selContratada.innerHTML = '<option value="todas">Todas as Empresas</option>' + contratadas.map(c => `<option value="${c}">${c}</option>`).join('');
    selBR.innerHTML = '<option value="todas">Todas</option>' + brs.map(b => `<option value="${b}">${b}</option>`).join('');
    if (contratadas.includes(ctrAtual)) selContratada.value = ctrAtual;
    if (brs.includes(brAtual)) selBR.value = brAtual;
}

function atualizarCardsUL() {
    const grid = document.getElementById('ulCardsGrid');
    const uls = [...new Set(dados.map(d => d.ul).filter(Boolean))].sort();
    const situacao = document.getElementById('filtroSituacao').value;
    const contratada = document.getElementById('filtroContratada').value;
    const br = document.getElementById('filtroBR').value;
    grid.innerHTML = '';

    uls.forEach(u => {
        const contratosUL = dados.filter(d => {
            if (d.ul !== u) return false;
            if (situacao !== 'todas' && !(d.situacao || '').toLowerCase().includes(situacao.toLowerCase())) return false;
            if (contratada !== 'todas' && d.contratada !== contratada) return false;
            if (br !== 'todas' && !(d.br || '').includes(br)) return false;
            return true;
        });

        if (contratosUL.length === 0) return;

        const totalContratos = contratosUL.length;
        const totalMensal = contratosUL.reduce((acc, d) => acc + MESES.reduce((s, m) => s + ((d.mensal && d.mensal[m]) || 0), 0), 0);
        const totalLoa = contratosUL.reduce((acc, d) => acc + (d.loa || 0), 0);
        const totalSaldoEmpenho = contratosUL.reduce((acc, d) => acc + (d.saldoEmpenho || 0), 0);
        const totalRap = contratosUL.reduce((acc, d) => acc + (d.rap || 0), 0);

        const card = document.createElement('div');
        card.className = `ul-card ${ulAtivaFiltro === u ? 'ativo' : ''}`;
        card.onclick = () => { ulAtivaFiltro = (ulAtivaFiltro === u) ? 'todas' : u; aplicarFiltros(); };

        card.innerHTML = `
            <div class="ul-card-title">${u} <span>${totalContratos} contratos</span></div>
            <div class="ul-card-info">
                <div class="ul-info-row"><span>Total Mensal:</span><strong>${formatarMoedaCurta(totalMensal)}</strong></div>
                <div class="ul-info-row"><span>LOA:</span><strong>${formatarMoedaCurta(totalLoa)}</strong></div>
                <div class="ul-info-row"><span>RAP:</span><strong>${formatarMoedaCurta(totalRap)}</strong></div>
            </div>
            <div class="ul-info-saldo"><div class="label">Saldo Empenho</div><div class="value">${formatarMoedaCurta(totalSaldoEmpenho)}</div></div>
        `;
        grid.appendChild(card);
    });
}

function aplicarFiltros() {
    const situacao = document.getElementById('filtroSituacao').value;
    const contratada = document.getElementById('filtroContratada').value;
    const br = document.getElementById('filtroBR').value;

    dadosFiltrados = dados.filter(item => {
        if (ulAtivaFiltro !== 'todas' && item.ul !== ulAtivaFiltro) return false;
        if (situacao !== 'todas' && !(item.situacao || '').toLowerCase().includes(situacao.toLowerCase())) return false;
        if (contratada !== 'todas' && item.contratada !== contratada) return false;
        if (br !== 'todas' && !(item.br || '').includes(br)) return false;
        return true;
    });

    atualizarCardsUL();
    renderizarTabelaContratos();
    renderizarTabelaMensal();
    atualizarKPI();
}

function limparFiltros() {
    ulAtivaFiltro = 'todas';
    document.getElementById('filtroSituacao').value = 'todas';
    document.getElementById('filtroContratada').value = 'todas';
    document.getElementById('filtroBR').value = 'todas';
    aplicarFiltros();
}

// ============================================================
// 📊 KPI
// ============================================================
function atualizarKPI() {
    const baseCalculo = dadosFiltrados;
    const totalContratos = baseCalculo.length;
    const totalLoa = baseCalculo.reduce((a, d) => a + (d.loa || 0), 0);
    const totalRap = baseCalculo.reduce((a, d) => a + (d.rap || 0), 0);
    const totalSaldoEmpenho = baseCalculo.reduce((a, d) => a + (d.saldoEmpenho || 0), 0);
    const totalMensal = baseCalculo.reduce((a, d) => a + MESES.reduce((s, m) => s + ((d.mensal && d.mensal[m]) || 0), 0), 0);

    document.getElementById('totalSelecionado').innerHTML = `
        <span>📋 Contratos: <strong>${totalContratos}</strong></span>
        <span>💵 LOA: <strong>${formatarMoeda(totalLoa)}</strong></span>
        <span>💰 RAP: <strong>${formatarMoeda(totalRap)}</strong></span>
        <span>📅 Mensal: <strong>${formatarMoeda(totalMensal)}</strong></span>
        <span style="color:#4ade80;">🛡️ Saldo Empenho: <strong>${formatarMoeda(totalSaldoEmpenho)}</strong></span>
    `;
}

// ============================================================
// 📊 TABELA CONTRATOS
// ============================================================
function renderizarTabelaContratos() {
    const tbody = document.getElementById('tabelaContratosBody');
    tbody.innerHTML = '';

    if (dadosFiltrados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="18" style="text-align:center;padding:20px;">Nenhum registro</td></tr>';
        atualizarBotaoExclusao();
        return;
    }

    dadosFiltrados.forEach(d => {
        const row = tbody.insertRow();
        const editClass = modoEdicao ? 'editavel' : '';
        const readonly = modoEdicao ? '' : 'readonly';

        const cCheck = row.insertCell(0);
        cCheck.innerHTML = `<input type="checkbox" class="checkbox-contrato" value="${d.id}" ${selecionadosIds.has(d.id) ? 'checked' : ''} onchange="toggleSelecaoContrato('${d.id}', this.checked)">`;

        row.insertCell(1).innerHTML = d.ul || '-';
        row.insertCell(2).innerHTML = d.contrato || '-';
        row.insertCell(3).innerHTML = d.br || '-';
        row.insertCell(4).innerHTML = d.servico || '-';
        row.insertCell(5).innerHTML = d.contratada || '-';

        const cSit = row.insertCell(6);
        if (modoEdicao) {
            cSit.innerHTML = `<select onchange="atualizarCampo('${d.id}', 'situacao', this.value)">
                <option value="Ativo" ${d.situacao === 'Ativo' ? 'selected' : ''}>Ativo</option>
                <option value="Concluído" ${d.situacao === 'Concluído' ? 'selected' : ''}>Concluído</option>
                <option value="Aguardando" ${(d.situacao || '').includes('Aguardando') ? 'selected' : ''}>Aguardando</option>
                <option value="Cancelado" ${d.situacao === 'Cancelado' ? 'selected' : ''}>Cancelado</option>
                <option value="" ${!d.situacao ? 'selected' : ''}>-</option>
            </select>`;
        } else {
            cSit.innerHTML = `<span class="badge ${getBadgeClass(d.situacao)}">${d.situacao || 'N/I'}</span>`;
        }

        row.insertCell(7).innerHTML = `<input type="text" value="${d.vigencia || ''}" ${readonly} class="${editClass}" onchange="atualizarCampo('${d.id}','vigencia',this.value)" style="width:90px;" placeholder="DD/MM/AAAA">`;
        row.insertCell(8).innerHTML = `<input type="number" value="${d.ultimaMedLancada || 0}" ${readonly} class="${editClass}" onchange="atualizarCampoNumInt('${d.id}','ultimaMedLancada',this.value)" style="width:70px;">`;
        row.insertCell(9).innerHTML = `<input type="text" value="${d.mesUltimaMedLancada || ''}" ${readonly} class="${editClass}" onchange="atualizarCampo('${d.id}','mesUltimaMedLancada',this.value)" style="width:90px;" placeholder="jun/26">`;
        row.insertCell(10).innerHTML = `<input type="number" value="${d.ultimaMedPaga || 0}" ${readonly} class="${editClass}" onchange="atualizarCampoNumInt('${d.id}','ultimaMedPaga',this.value)" style="width:70px;">`;
        row.insertCell(11).innerHTML = `<input type="text" value="${d.mesUltimaMedPaga || ''}" ${readonly} class="${editClass}" onchange="atualizarCampo('${d.id}','mesUltimaMedPaga',this.value)" style="width:90px;" placeholder="mai/26">`;

        const cSaldoEmpenho = row.insertCell(12);
        cSaldoEmpenho.style.background = '#e8f5e9'; cSaldoEmpenho.style.fontWeight = 'bold'; cSaldoEmpenho.style.color = '#1b5e20';
        cSaldoEmpenho.innerHTML = `<input type="text" value="${formatarNumeroInput(d.saldoEmpenho)}" ${readonly} class="${editClass}" onchange="atualizarCampoNum('${d.id}','saldoEmpenho',this.value)" style="width:110px; font-weight:bold; color:#1b5e20; border-color:#81c784;">`;

        row.insertCell(13).innerHTML = `<input type="text" value="${formatarNumeroInput(d.rap)}" ${readonly} class="${editClass}" onchange="atualizarCampoNum('${d.id}','rap',this.value)" style="width:100px;">`;
        row.insertCell(14).innerHTML = `<input type="text" value="${formatarNumeroInput(d.loa)}" ${readonly} class="${editClass}" onchange="atualizarCampoNum('${d.id}','loa',this.value)" style="width:110px;">`;
        row.insertCell(15).innerHTML = `<input type="text" value="${formatarNumeroInput(d.remanejMenos)}" ${readonly} class="${editClass}" onchange="atualizarCampoNum('${d.id}','remanejMenos',this.value)" style="width:100px;">`;
        row.insertCell(16).innerHTML = `<input type="text" value="${formatarNumeroInput(d.remanejMais)}" ${readonly} class="${editClass}" onchange="atualizarCampoNum('${d.id}','remanejMais',this.value)" style="width:100px;">`;

        const cAcoes = row.insertCell(17);
        cAcoes.innerHTML = `
            <button class="acao-btn acao-btn-view" onclick="verDetalhes('${d.id}')" title="Ver">👁️</button>
            <button class="acao-btn acao-btn-del" onclick="excluirContrato('${d.id}')" title="Excluir">🗑️</button>
        `;
    });

    atualizarBotaoExclusao();
    const checkAll = document.getElementById('checkAll');
    const allChecked = dadosFiltrados.length > 0 && dadosFiltrados.every(d => selecionadosIds.has(d.id));
    checkAll.checked = allChecked;
}

// ============================================================
// 📅 TABELA MENSAL
// ============================================================
function renderizarTabelaMensal() {
    const thead = document.getElementById('tabelaMensalHead');
    const tbody = document.getElementById('tabelaMensalBody');

    let headHtml = '<th class="contrato-cell">Contrato</th><th>UL</th>';
    MESES.forEach(m => {
        const [ano, mes] = m.split('-');
        headHtml += `<th class="mes-col">${mes}/${ano.slice(2)}</th>`;
    });
    headHtml += '<th>Total</th>';
    thead.innerHTML = headHtml;

    tbody.innerHTML = '';
    if (dadosFiltrados.length === 0) {
        tbody.innerHTML = `<tr><td colspan="${MESES.length + 3}" style="text-align:center;padding:20px;">Nenhum registro</td></tr>`;
        return;
    }

    dadosFiltrados.forEach(d => {
        const row = tbody.insertRow();
        const readonly = modoEdicao ? '' : 'readonly';

        const cContrato = row.insertCell(0);
        cContrato.className = 'contrato-cell';
        cContrato.innerHTML = `<strong>${d.contrato}</strong><br><small style="color:#64748b;">${(d.contratada || '').substring(0, 25)}${(d.contratada || '').length > 25 ? '...' : ''}</small>`;

        row.insertCell(1).innerHTML = d.ul || '-';

        let total = 0;
        MESES.forEach(m => {
            const valor = (d.mensal && d.mensal[m]) || 0;
            total += valor;
            const cell = row.insertCell();
            cell.innerHTML = `<input type="text" value="${formatarNumeroInput(valor)}" ${readonly} onchange="atualizarMensal('${d.id}','${m}',this.value)" placeholder="0,00">`;
        });

        const cTotal = row.insertCell();
        cTotal.className = 'total-col';
        cTotal.innerHTML = formatarMoeda(total);
    });
}

// ============================================================
// ✏️ ATUALIZAÇÕES
// ============================================================
function atualizarCampo(id, campo, valor) {
    const item = dados.find(d => d.id === id); if (!item) return;
    item[campo] = valor; salvarItemFirebase(id);
}
function atualizarCampoNum(id, campo, valor) {
    const item = dados.find(d => d.id === id); if (!item) return;
    item[campo] = parseNumero(valor); salvarItemFirebase(id); aplicarFiltros();
}
function atualizarCampoNumInt(id, campo, valor) {
    const item = dados.find(d => d.id === id); if (!item) return;
    item[campo] = parseInt(valor) || 0; salvarItemFirebase(id);
}
function atualizarMensal(id, mes, valor) {
    const item = dados.find(d => d.id === id); if (!item) return;
    if (!item.mensal) item.mensal = {};
    item.mensal[mes] = parseNumero(valor);

    const total = MESES.reduce((s, m) => s + ((item.mensal && item.mensal[m]) || 0), 0);
    const rows = document.querySelectorAll('#tabelaMensalBody tr');
    rows.forEach(row => {
        const firstCell = row.cells[0];
        if (firstCell && firstCell.textContent.includes(item.contrato)) {
            const lastCell = row.cells[row.cells.length - 1];
            lastCell.innerHTML = formatarMoeda(total);
        }
    });

    atualizarKPI(); atualizarCardsUL(); salvarItemFirebase(id);
}

// ============================================================
// 👁️ DETALHES
// ============================================================
function verDetalhes(id) {
    const item = dados.find(d => d.id === id); if (!item) return;
    const total = MESES.reduce((s, m) => s + ((item.mensal && item.mensal[m]) || 0), 0);
    alert(
        `📋 CONTRATO: ${item.contrato}\n🏢 UL: ${item.ul}\n🛣️ BR: ${item.br}\n📝 Serviço: ${item.servico}\n🏗️ Contratada: ${item.contratada}\n📌 Situação: ${item.situacao || 'N/I'}\n📅 Vigência: ${item.vigencia || '-'}\n🔢 Últ. Med. Lançada: ${item.ultimaMedLancada || '-'} (${item.mesUltimaMedLancada || '-'})\n🔢 Últ. Med. Paga: ${item.ultimaMedPaga || '-'} (${item.mesUltimaMedPaga || '-'})\n🛡️ Saldo Empenho: ${formatarMoeda(item.saldoEmpenho)}\n💰 RAP: ${formatarMoeda(item.rap)}\n💰 LOA: ${formatarMoeda(item.loa)}\n📊 Total Mensal: ${formatarMoeda(total)}`
    );
}

// ============================================================
// 🔄 RESETAR / ESVAZIAR
// ============================================================
function resetarDados() {
    if (!confirm('⚠️ RESETAR todos os dados?\n\nIsso apagará TODAS as alterações e contratos adicionados.')) return;
    mostrarStatus('🔄 Resetando...', 'info');
    database.ref(REF_CONTRATOS).remove()
        .then(() => {
            dados = DADOS_BASE.map(d => ({ ...d, mensal: MESES.reduce((acc, m) => { acc[m] = 0; return acc; }, {}) }));
            dadosCarregados = true; salvarDadosFirebase(true); popularFiltros(); aplicarFiltros();
        }).catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
}

function esvaziarDados() {
    if (!confirm('⚠️ APAGAR TODOS OS CONTRATOS?\n\nIsso remove todos os contratos do banco. Não dá para desfazer.')) return;
    mostrarStatus('🧹 Apagando...', 'info');
    const updates = {};
    updates[REF_CONTRATOS] = null; updates[`${REF_CONFIG}/baseVazia`] = true;
    updates[`${REF_CONFIG}/ultimaAtualizacao`] = new Date().toISOString();
    updates[`${REF_CONFIG}/usuario`] = auth.currentUser ? auth.currentUser.email : 'desconhecido';
    database.ref().update(updates)
        .then(() => { dados = []; dadosCarregados = true; popularFiltros(); aplicarFiltros(); mostrarStatus('✅ Tudo apagado!'); })
        .catch(err => mostrarStatus('❌ Erro: ' + err.message, 'erro'));
}

// ============================================================
// ⌨️ ATALHOS E CLIQUES FORA
// ============================================================
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        fecharModal(); fecharModalAdicionar(); fecharModalUpload();
    }
});
document.addEventListener('click', function(e) {
    if (e.target.classList.contains('modal-overlay')) {
        if (e.target.id === 'modalUpload') fecharModalUpload();
        else if (e.target.id === 'modalAdicionar') fecharModalAdicionar();
        else if (e.target.id === 'modalConfirm') fecharModal();
    }
});

async function abrirAppParaUsuario(user) {
    if (appAbertoPara === user.uid) return;   // já aberto, ignora evento duplicado
    appAbertoPara = user.uid;

    const username = user.email.split('@')[0];

    document.getElementById('loginOverlay').style.display = 'none';
    document.getElementById('appContainer').style.display = 'none';
    document.getElementById('modalTrocaSenha').style.display = 'none';

    try {
        // dev1: acesso direto, sem verificação de troca de senha
        if (username !== 'dev1') {
            const snapshot = await database.ref(`${REF_USUARIOS}/${username}`).once('value');
            const userData = snapshot.val();

            if (!userData) {
                await database.ref(`${REF_USUARIOS}/${username}`).set({
                    precisaTrocarSenha: true,
                    primeiroAcesso: new Date().toISOString()
                });
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
        if (!dadosCarregados) carregarDados();

    } catch (error) {
        console.error('Erro ao verificar acesso:', error);
        document.getElementById('appContainer').style.display = 'block';
        if (!dadosCarregados) carregarDados();
    }
}
