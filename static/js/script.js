// Scripts JavaScript para o sistema Extra PDRLOG

document.addEventListener('DOMContentLoaded', function() {
    // Atualizar contadores do sidebar
    atualizarContadoresSidebar();
    
    // Configurar tooltips do Bootstrap
    var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'))
    var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl)
    });
});

function atualizarContadoresSidebar() {
    // Fazer requisição para API de estatísticas
    fetch('/api/dashboard/stats')
        .then(response => response.json())
        .then(data => {
            // Atualizar contadores no sidebar
            const solicitacoesPendentes = document.getElementById('status-solicitacoes-pendentes');
            const ncsAbertas = document.getElementById('status-nc-abertas');
            
            if (solicitacoesPendentes) {
                solicitacoesPendentes.textContent = data.solicitacoes_aguardando || 0;
            }
            
            if (ncsAbertas) {
                // Para NCs abertas, usar um cálculo baseado no status
                ncsAbertas.textContent = data.ncs_pendentes || 0;
            }
        })
        .catch(error => {
            console.error('Erro ao carregar estatísticas:', error);
        });
}

// Função para formatar moeda
function formatarMoeda(valor) {
    if (!valor) return '0,00';
    
    // Converte para número
    let numero = parseFloat(valor.toString().replace(/[^\d,-]/g, '').replace(',', '.'));
    if (isNaN(numero)) numero = 0;
    
    // Formata como moeda brasileira
    return numero.toLocaleString('pt-BR', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

// Função para converter moeda para float
function moedaParaFloat(moeda) {
    if (!moeda) return 0;
    let valor = moeda.toString();
    if (valor.includes(',')) {
        let partes = valor.split(',');
        partes[0] = partes[0].replace(/\./g, '');
        valor = partes.join('.');
    }
    return parseFloat(valor);
}

// Função para mostrar loading
function mostrarLoading(mensagem = 'Processando...') {
    // Implementar overlay de loading se necessário
    console.log(mensagem);
}

// Função para validar formulários
function validarFormularioSolicitacao() {
    const form = document.getElementById('formSolicitacao');
    if (!form) return true;
    
    const oms = form.querySelectorAll('input[name="om[]"]');
    let temOmValida = false;
    
    for (let om of oms) {
        if (om.value && om.value.trim()) {
            temOmValida = true;
            break;
        }
    }
    
    if (!temOmValida) {
        alert('É necessário informar pelo menos uma OM!');
        return false;
    }
    
    return true;
}

// Exportar funções para uso global
window.formatarMoeda = formatarMoeda;
window.moedaParaFloat = moedaParaFloat;
window.mostrarLoading = mostrarLoading;
window.validarFormularioSolicitacao = validarFormularioSolicitacao;