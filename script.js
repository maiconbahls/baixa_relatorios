// Global states for report data
let globalReportData = null;
let filteredValues = null;

// On document load
document.addEventListener("DOMContentLoaded", () => {
    checkPortalSession();
    restoreInterviewCredentials();
    setupEventListeners();
});

// Setup application event listeners
function setupEventListeners() {
    // Portal Login Form
    const loginForm = document.getElementById("portal-login-form");
    if (loginForm) {
        loginForm.addEventListener("submit", handlePortalLogin);
    }

    // Interview Download Button
    const btnDownload = document.getElementById("btn-download");
    if (btnDownload) {
        btnDownload.addEventListener("click", handleDownloadReport);
    }

    // Portal Logout Button
    const btnLogout = document.getElementById("btn-logout");
    if (btnLogout) {
        btnLogout.addEventListener("click", handlePortalLogout);
    }
}

// ----------------------------------------------------
// 1. PORTAL LOGIN SYSTEM
// ----------------------------------------------------

function handlePortalLogin(e) {
    e.preventDefault();
    
    const usernameInput = document.getElementById("portal-username").value.trim().toLowerCase();
    const passwordInput = document.getElementById("portal-password").value.trim().toLowerCase();
    const errorAlert = document.getElementById("login-error");
    
    // Accept 'gestão' or 'gestao' for user-friendliness
    const isValidUser = (usernameInput === "gestão" || usernameInput === "gestao");
    const isValidPass = (passwordInput === "gestão" || passwordInput === "gestao");
    
    if (isValidUser && isValidPass) {
        // Save session in sessionStorage (expires when tab is closed)
        sessionStorage.setItem("portal_session_active", "true");
        errorAlert.classList.add("hidden");
        
        // Transition screens
        document.getElementById("login-container").classList.remove("active");
        document.getElementById("dashboard-container").classList.remove("hidden");
        
        showToast("Bem-vindo!", "Acesso ao Portal de Gestão realizado com sucesso.", "success");
    } else {
        errorAlert.classList.remove("hidden");
    }
}

function checkPortalSession() {
    const isSessionActive = sessionStorage.getItem("portal_session_active");
    if (isSessionActive === "true") {
        document.getElementById("login-container").classList.remove("active");
        document.getElementById("dashboard-container").classList.remove("hidden");
    } else {
        document.getElementById("login-container").classList.add("active");
        document.getElementById("dashboard-container").classList.add("hidden");
    }
}

function handlePortalLogout() {
    sessionStorage.removeItem("portal_session_active");
    
    // Clear preview table states
    globalReportData = null;
    filteredValues = null;
    
    // Reset table elements
    document.getElementById("data-preview-table").classList.add("hidden");
    document.getElementById("table-placeholder").classList.remove("hidden");
    document.getElementById("preview-controls").classList.add("hidden");
    document.getElementById("preview-subtitle").innerText = "Nenhum dado carregado ainda.";
    document.getElementById("terminal-log").innerHTML = '<div class="log-line text-muted">Aguardando início do processo...</div>';
    
    // Clear portal inputs
    document.getElementById("portal-username").value = "";
    document.getElementById("portal-password").value = "";
    
    checkPortalSession();
    showToast("Sessão Encerrada", "Você saiu do portal com segurança.", "info");
}

// Helper to show password
function togglePasswordVisibility(inputId, iconElement) {
    const input = document.getElementById(inputId);
    if (input.type === "password") {
        input.type = "text";
        iconElement.classList.replace("fa-eye", "fa-eye-slash");
    } else {
        input.type = "password";
        iconElement.classList.replace("fa-eye-slash", "fa-eye");
    }
}

// ----------------------------------------------------
// 2. CREDENTIALS MANAGEMENT
// ----------------------------------------------------

function restoreInterviewCredentials() {
    const savedUser = localStorage.getItem("app_interview_user");
    const savedPass = localStorage.getItem("app_interview_pass");
    const saveCheckbox = localStorage.getItem("app_interview_save");
    
    if (savedUser) {
        document.getElementById("interview-username").value = savedUser;
    }
    if (savedPass) {
        document.getElementById("interview-password").value = savedPass;
    }
    if (saveCheckbox === "false") {
        document.getElementById("save-credentials").checked = false;
    }
}

function saveInterviewCredentials(usuario, senha, saveChecked) {
    if (saveChecked) {
        localStorage.setItem("app_interview_user", usuario);
        localStorage.setItem("app_interview_pass", senha);
        localStorage.setItem("app_interview_save", "true");
    } else {
        localStorage.removeItem("app_interview_user");
        localStorage.removeItem("app_interview_pass");
        localStorage.setItem("app_interview_save", "false");
    }
}

// Predefined reports mapping for friendly names and filenames
const PREDEFINED_REPORTS = {
    '132': { id: '132', name: 'Colaborador Senior', file: 'Colaborador_Senior' },
    '310': { id: '310', name: 'Colaboradores Grupo - Geral', file: 'Colaboradores_Grupo_Geral' },
    '316': { id: '316', name: 'Colaboradores Senior - MS', file: 'Colaboradores_Senior_MS' },
    '372': { id: '372', name: 'Colaboradores com Posto de Trabalho', file: 'Colaboradores_Posto_Trabalho' }
};

// Handle report selector dropdown change
function handleReportSelectChange(selectElement) {
    const customWrapper = document.getElementById("custom-report-id-wrapper");
    const btnText = document.querySelector(".main-btn-text");
    const btnSub = document.querySelector(".main-btn-sub");
    
    if (selectElement.value === "custom") {
        customWrapper.classList.remove("hidden");
        document.getElementById("custom-report-id").focus();
        btnText.innerText = "Iniciar Integração";
        btnSub.innerText = "Efetuar login e obter dados";
    } else {
        customWrapper.classList.add("hidden");
        if (selectElement.value === "all") {
            btnText.innerText = "Iniciar Integração (Lote)";
            btnSub.innerText = "Baixar 4 relatórios em lote";
        } else {
            btnText.innerText = "Iniciar Integração";
            btnSub.innerText = "Efetuar login e obter dados";
        }
    }
}

// Initialize the dropdown helper text on page load
document.addEventListener("DOMContentLoaded", () => {
    const selectElement = document.getElementById("report-id-select");
    if (selectElement) {
        handleReportSelectChange(selectElement);
    }
});

// ----------------------------------------------------
// 3. LOGGING / TERMINAL CONSOLE
// ----------------------------------------------------

function clearLog() {
    const consoleLog = document.getElementById("terminal-log");
    consoleLog.innerHTML = "";
}

function addLog(message, type = "normal") {
    const consoleLog = document.getElementById("terminal-log");
    const timestamp = new Date().toLocaleTimeString("pt-BR");
    
    const line = document.createElement("div");
    line.className = `log-line ${type}`;
    line.innerHTML = `<span class="text-muted">[${timestamp}]</span> ${message}`;
    
    consoleLog.appendChild(line);
    consoleLog.scrollTop = consoleLog.scrollHeight;
}

// ----------------------------------------------------
// 4. REPORT DOWNLOAD & INTEGRATION
// ----------------------------------------------------

async function handleDownloadReport() {
    const usuario = document.getElementById("interview-username").value.trim();
    const senha = document.getElementById("interview-password").value;
    const saveChecked = document.getElementById("save-credentials").checked;
    const reportSelect = document.getElementById("report-id-select").value;

    if (!usuario || !senha) {
        showToast("Erro de entrada", "Preencha o usuário e a senha do Interview.", "error");
        addLog("Erro: Credenciais do Interview não foram preenchidas.", "error");
        return;
    }

    // Save credentials to localStorage if selected
    saveInterviewCredentials(usuario, senha, saveChecked);

    // Update UI loading state
    const btnDownload = document.getElementById("btn-download");
    btnDownload.classList.add("loading");
    btnDownload.disabled = true;
    
    clearLog();

    if (reportSelect === "all") {
        addLog("Iniciando download em LOTE (4 relatórios)...", "info");
        const listToDownload = Object.values(PREDEFINED_REPORTS);
        let successCount = 0;
        
        for (let i = 0; i < listToDownload.length; i++) {
            const report = listToDownload[i];
            addLog(`--------------------------------------------------`, "normal");
            addLog(`[LOTE ${i+1}/${listToDownload.length}] Baixando: ${report.name} (ID ${report.id})...`, "info");
            addLog("Autenticando e extraindo dados...", "normal");
            
            try {
                const response = await fetch("/api/baixar", {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({ usuario, senha, id: report.id })
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.error || `Erro HTTP: ${response.status}`);
                }

                addLog(`[LOTE] Sucesso! Registros recuperados: ${data.valores.length} linhas.`, "success");
                addLog(`[LOTE] Gerando planilha Excel...`, "normal");
                
                const dateStr = new Date().toLocaleDateString("pt-BR").replace(/\//g, "-");
                const fileName = `${report.file}_${dateStr}.xlsx`;
                
                generateAndDownloadExcel(data.colunas, data.valores, fileName, report.name);
                addLog(`[LOTE] Download concluído: ${fileName}`, "success");
                
                // Show last processed report in preview grid
                globalReportData = data;
                filteredValues = data.valores;
                renderPreviewTable(data.colunas, data.valores);
                
                successCount++;
            } catch (err) {
                addLog(`[LOTE] Erro no relatório "${report.name}": ${err.message}`, "error");
            }
            
            // Wait a brief moment between downloads to prevent overloading the server
            await new Promise(resolve => setTimeout(resolve, 800));
        }
        
        addLog(`--------------------------------------------------`, "normal");
        addLog(`Lote finalizado! Relatórios baixados com sucesso: ${successCount} de ${listToDownload.length}.`, "info");
        
        if (successCount === listToDownload.length) {
            showToast("Lote Concluído!", "Todos os 4 relatórios foram baixados.", "success");
        } else if (successCount > 0) {
            showToast("Lote Parcial!", `${successCount} relatórios foram salvos, alguns falharam.`, "info");
        } else {
            showToast("Falha no Lote", "Não foi possível baixar nenhum relatório.", "error");
        }
        
    } else {
        // Single Report Download
        let reportId = reportSelect;
        let reportName = "Relatório";
        let fileBaseName = "Relatorio";
        
        if (reportSelect === "custom") {
            reportId = document.getElementById("custom-report-id").value.trim();
            if (!reportId) {
                showToast("Aviso", "Por favor, preencha o ID do relatório customizado.", "error");
                btnDownload.classList.remove("loading");
                btnDownload.disabled = false;
                return;
            }
            reportName = `Relatório ID ${reportId}`;
            fileBaseName = `Relatorio_ID_${reportId}`;
        } else if (PREDEFINED_REPORTS[reportSelect]) {
            reportName = PREDEFINED_REPORTS[reportSelect].name;
            fileBaseName = PREDEFINED_REPORTS[reportSelect].file;
        }

        addLog(`Iniciando download do relatório: ${reportName} (ID: ${reportId})...`, "info");
        addLog("Conectando ao proxy e efetuando login...", "normal");

        try {
            const response = await fetch("/api/baixar", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ usuario, senha, id: reportId })
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.error || `Erro HTTP: ${response.status}`);
            }

            addLog("Autenticação e carregamento de dados concluídos!", "success");
            addLog(`Total de registros recuperados: ${data.valores.length} linhas.`, "info");
            
            globalReportData = data;
            filteredValues = data.valores;

            // Render preview table
            renderPreviewTable(data.colunas, data.valores);
            
            // Trigger Excel download
            addLog("Gerando arquivo Excel (.xlsx)...", "info");
            const dateStr = new Date().toLocaleDateString("pt-BR").replace(/\//g, "-");
            const fileName = `${fileBaseName}_${dateStr}.xlsx`;
            generateAndDownloadExcel(data.colunas, data.valores, fileName, reportName);
            
            addLog(`Arquivo baixado com sucesso: ${fileName}`, "success");
            showToast("Sucesso!", "Relatório baixado.", "success");

        } catch (err) {
            addLog(`Erro durante a integração: ${err.message}`, "error");
            showToast("Erro na integração", err.message, "error");
            
            globalReportData = null;
            filteredValues = null;
            document.getElementById("data-preview-table").classList.add("hidden");
            document.getElementById("table-placeholder").classList.remove("hidden");
            document.getElementById("preview-controls").classList.add("hidden");
            document.getElementById("preview-subtitle").innerText = "Falha ao carregar dados.";
        }
    }

    // Restore button state
    btnDownload.classList.remove("loading");
    btnDownload.disabled = false;
}

// ----------------------------------------------------
// 5. DATA PREVIEW TABLE & FILTERING
// ----------------------------------------------------

function renderPreviewTable(colunas, valores) {
    const tableHeaders = document.getElementById("table-headers");
    const tableBody = document.getElementById("table-body");
    const tablePlaceholder = document.getElementById("table-placeholder");
    const dataTable = document.getElementById("data-preview-table");
    const previewControls = document.getElementById("preview-controls");
    const previewSubtitle = document.getElementById("preview-subtitle");

    // Clear previous elements
    tableHeaders.innerHTML = "";
    tableBody.innerHTML = "";

    // Set headers
    colunas.forEach(col => {
        const th = document.createElement("th");
        th.innerText = col.title || col;
        tableHeaders.appendChild(th);
    });

    // Populate rows (show max 10 rows for preview performance)
    const previewRows = valores.slice(0, 10);
    previewRows.forEach(row => {
        const tr = document.createElement("tr");
        row.forEach(val => {
            const td = document.createElement("td");
            td.innerText = val === null || val === undefined ? "" : val;
            td.title = td.innerText; // Tooltip for full content
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });

    // Switch view
    tablePlaceholder.classList.add("hidden");
    dataTable.classList.remove("hidden");
    previewControls.classList.remove("hidden");
    
    // Reset search filter input
    document.getElementById("preview-search").value = "";

    // Update subtitles
    previewSubtitle.innerText = `Amostra de 10 de ${valores.length} linhas encontradas.`;
}

function filterPreviewData(searchTerm) {
    if (!globalReportData) return;
    
    const query = searchTerm.toLowerCase().trim();
    
    // Filter rows that contain query in any cell
    filteredValues = globalReportData.valores.filter(row => {
        return row.some(cell => {
            const cellStr = (cell === null || cell === undefined ? "" : cell).toString().toLowerCase();
            return cellStr.includes(query);
        });
    });

    // Rerender table matching filtered values
    const tableBody = document.getElementById("table-body");
    tableBody.innerHTML = "";

    const previewRows = filteredValues.slice(0, 10);
    previewRows.forEach(row => {
        const tr = document.createElement("tr");
        row.forEach(val => {
            const td = document.createElement("td");
            td.innerText = val === null || val === undefined ? "" : val;
            td.title = td.innerText;
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });

    // Update subtitle
    const previewSubtitle = document.getElementById("preview-subtitle");
    if (query === "") {
        previewSubtitle.innerText = `Amostra de 10 de ${globalReportData.valores.length} linhas encontradas.`;
    } else {
        previewSubtitle.innerText = `Amostra de 10 de ${filteredValues.length} linhas filtradas (de ${globalReportData.valores.length} totais).`;
    }
}

// ----------------------------------------------------
// 6. CLIENT-SIDE EXCEL EXPORT (SHEETJS)
// ----------------------------------------------------

// Generalized helper function to export data to XLSX
function generateAndDownloadExcel(colunas, valores, filename, sheetname) {
    const headerRow = colunas.map(col => col.title || col);
    const dataToExport = [headerRow, ...valores];
    const worksheet = XLSX.utils.aoa_to_sheet(dataToExport);
    
    // Auto fit columns width
    const colWidths = headerRow.map((colName, index) => {
        let maxLength = colName.toString().length;
        const sampleRows = valores.slice(0, 100);
        sampleRows.forEach(row => {
            const val = row[index];
            if (val !== null && val !== undefined) {
                maxLength = Math.max(maxLength, val.toString().length);
            }
        });
        return { wch: Math.min(maxLength + 3, 35) }; // Limit to max 35 chars
    });
    worksheet['!cols'] = colWidths;
    
    const workbook = XLSX.utils.book_new();
    // Excel sheet name must be less than 31 characters
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetname.substring(0, 30));
    XLSX.writeFile(workbook, filename);
}

function triggerExcelDownload() {
    if (!globalReportData || filteredValues.length === 0) {
        showToast("Sem dados", "Não há dados para exportar.", "error");
        return;
    }

    try {
        const reportSelect = document.getElementById("report-id-select").value;
        const dateStr = new Date().toLocaleDateString("pt-BR").replace(/\//g, "-");
        
        let fileBase = "Relatorio";
        let sheetName = "Relatório";
        
        if (PREDEFINED_REPORTS[reportSelect]) {
            fileBase = PREDEFINED_REPORTS[reportSelect].file;
            sheetName = PREDEFINED_REPORTS[reportSelect].name;
        } else if (reportSelect === "custom") {
            const customId = document.getElementById("custom-report-id").value.trim();
            fileBase = `Relatorio_ID_${customId}`;
            sheetName = `ID ${customId}`;
        }
        
        const fileName = `${fileBase}_${dateStr}.xlsx`;
        generateAndDownloadExcel(globalReportData.colunas, filteredValues, fileName, sheetName);
        
        addLog(`Arquivo baixado com sucesso: ${fileName}`, "success");
        showToast("Sucesso!", "Excel salvo localmente.", "success");
    } catch (e) {
        addLog(`Erro ao gerar Excel: ${e.message}`, "error");
        showToast("Erro ao exportar", e.message, "error");
    }
}

// ----------------------------------------------------
// 7. TOAST NOTIFICATION UTILITY
// ----------------------------------------------------

function showToast(title, message, type = "success") {
    const toast = document.getElementById("toast-notification");
    const toastTitle = document.getElementById("toast-title");
    const toastMsg = document.getElementById("toast-message");
    const toastIcon = document.getElementById("toast-icon");
    
    toastTitle.innerText = title;
    toastMsg.innerText = message;
    
    // Set icon depending on status type
    if (type === "success") {
        toastIcon.className = "fa-solid fa-circle-check toast-icon";
        toastIcon.style.color = "var(--success)";
    } else if (type === "error") {
        toastIcon.className = "fa-solid fa-circle-exclamation toast-icon";
        toastIcon.style.color = "var(--danger)";
    } else {
        toastIcon.className = "fa-solid fa-circle-info toast-icon";
        toastIcon.style.color = "var(--accent-teal)";
    }
    
    toast.classList.remove("hidden");
    
    // Auto close toast after 5 seconds
    setTimeout(() => {
        closeToast();
    }, 5000);
}

function closeToast() {
    const toast = document.getElementById("toast-notification");
    if (toast) {
        toast.classList.add("hidden");
    }
}
