$(document).ready(function() {
    let table = null, allData = [], nameConflicts = new Set(), dupLicUsers = new Set();
    const DEBOUNCE_DELAY = 350;
    let multiSearchDebounceTimer;

    const searchableColumnsConfig = [
        { index: 0, title: 'ID', dataProp: 'Id', useDropdown: false },
        { index: 1, title: 'Name', dataProp: 'DisplayName', useDropdown: false },
        { index: 2, title: 'Email', dataProp: 'Email', useDropdown: false },
        { index: 3, title: 'Job Title', dataProp: 'JobTitle', useDropdown: true },
        { index: 4, title: 'Location', dataProp: 'OfficeLocation', useDropdown: true },
        { index: 5, title: 'Phones', dataProp: 'BusinessPhones', useDropdown: false },
        { index: 6, title: 'Licenses', dataProp: 'Licenses', useDropdown: true }
    ];

    let uniqueFieldValues = {};
    const MAX_DROPDOWN_OPTIONS_DISPLAYED = 50;

    // --- Funções Utilitárias (loader, escape, etc.) ---
    function showLoader(message = 'Processing...') {
        let $overlay = $('#loadingOverlay');
        if ($overlay.length === 0 && $('body').length > 0) {
            $('body').append('<div id="loadingOverlay"><div class="loader-content"><p id="loaderMessageText"></p></div></div>');
            $overlay = $('#loadingOverlay');
        }
        $overlay.find('#loaderMessageText').text(message);
        $overlay.css('display', 'flex');
    }

    function hideLoader() {
        $('#loadingOverlay').hide();
    }

    const nameKey = u => `${u.DisplayName || ''}|||${u.OfficeLocation || ''}`;
    const escapeHtml = s => typeof s === 'string' ? s.replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[m]) : '';

    function escapeCsvValue(value, forceQuotesOnSemicolon = false) {
        if (value == null) return '';
        let stringValue = String(value);
        const regex = forceQuotesOnSemicolon ? /[,"\n\r;]/ : /[,"\n\r]/;
        if (regex.test(stringValue)) {
            stringValue = `"${stringValue.replace(/"/g, '""')}"`;
        }
        return stringValue;
    }
    // --- Fim Funções Utilitárias ---

    function renderAlerts() {
        const $alertPanel = $('#alertPanel').empty();
        if (nameConflicts.size) {
            $alertPanel.append(`<div class="alert-badge"><button id="filterNameConflicts" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;" class="underline"><i class="fas fa-users-slash" style="margin-right: 0.3rem;"></i>Name+Location Conflicts: ${nameConflicts.size}</button></div>`);
        }
        if (dupLicUsers.size) { // dupLicUsers agora é preenchido por findUsersWithDuplicateLicenses
            const usersToList = allData.filter(u => dupLicUsers.has(u.Id));
            let listHtml = '';
            const maxPreview = 10;
            usersToList.slice(0, maxPreview).forEach(u => {
                const licDetails = [];
                if (u.duplicateLicenseDetails) { // Usar a propriedade adicionada
                    for (const [licName, count] of Object.entries(u.duplicateLicenseDetails)) {
                        if (count > 1) licDetails.push(`${licName} (x${count})`);
                    }
                }
                const hasPaid = (u.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free') && !(l.LicenseName || '').toLowerCase().includes('trial'));
                listHtml += `<li>${escapeHtml(u.DisplayName)} (${escapeHtml(u.OfficeLocation || 'N/A')}): Duplicates: ${escapeHtml(licDetails.join(', ') || 'N/A')} — Has Paid: ${hasPaid ? 'Yes' : 'No'}</li>`;
            });
            if (usersToList.length > maxPreview) {
                listHtml += `<li>And ${usersToList.length - maxPreview} more user(s)...</li>`;
            }
            $alertPanel.append(`<div class="alert-badge"><span><i class="fas fa-copy" style="margin-right: 0.3rem;"></i>Users with Duplicate Licenses: ${dupLicUsers.size}</span> <button class="underline toggle-details" data-target="dupDetails" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;">Details</button></div><div id="dupDetails" class="alert-details"><ul>${listHtml}</ul></div>`);
        }

        $('#filterNameConflicts').off('click').on('click', function() { /* ... (mantido) ... */ });
        $('.toggle-details').off('click').on('click', function() { /* ... (mantido) ... */ });
    }


    function initTable(data) {
        allData = data;
        // Pré-processar para encontrar duplicatas uma vez no carregamento para o alerta
        const duplicateInfo = findUsersWithDuplicateLicenses(allData, { identifyOnly: true });
        dupLicUsers = new Set(duplicateInfo.map(u => u.Id)); // Atualiza o Set global para renderAlerts

        if ($('#licenseDatalist').length === 0) { $('body').append('<datalist id="licenseDatalist"></datalist>'); }
        const $licenseDatalist = $('#licenseDatalist').empty();
        if (uniqueFieldValues.Licenses) {
            uniqueFieldValues.Licenses.forEach(name => { $licenseDatalist.append($('<option>').attr('value', escapeHtml(name))); });
        }

        if (table) { $(table.table().node()).off('preDraw.dt draw.dt'); table.destroy(); $('#licenseTable').empty(); }

        table = $('#licenseTable').DataTable({
            data: allData,
            deferRender: true,
            pageLength: 25,
            orderCellsTop: true,
            columns: [
                { data: 'Id', title: 'ID', visible: false },
                { data: 'DisplayName', title: 'Name', visible: true },
                { data: 'Email', title: 'Email', visible: true },
                { data: 'JobTitle', title: 'Job Title', visible: true },
                { data: 'OfficeLocation', title: 'Location', visible: false },
                { data: 'BusinessPhones', title: 'Phones', visible: false, render: p => Array.isArray(p) ? p.join('; ') : (p || '') },
                {
                    data: 'Licenses', title: 'Licenses', visible: true,
                    render: function(licenses, type, row) {
                        let displayLicenses = Array.isArray(licenses) ? licenses.map(x => x.LicenseName || '').filter(name => name) : [];
                        if (row.duplicateLicenseDetails) { // Adiciona indicador visual de duplicatas
                            const dupIndicator = Object.entries(row.duplicateLicenseDetails)
                                .filter(([, count]) => count > 1)
                                .map(([name]) => `<span class="dup-marker" title="Duplicate: ${name}">⚠️</span>`)
                                .join(' ');
                            if (dupIndicator) {
                                return displayLicenses.join(', ') + ' ' + dupIndicator;
                            }
                        }
                        return displayLicenses.join(', ');
                    }
                }
            ],
            initComplete: function() {
                const api = this.api();
                $('#colContainer .col-vis').each(function() { /* ... (mantido) ... */ });
                $('.col-vis').off('change').on('change', function() { /* ... (mantido) ... */ });
                if ($('#multiSearchFields .multi-search-row').length > 0) { applyMultiSearch(); }
                updateAiFeatureStatus();
            },
            rowCallback: function(row, data) {
                let classes = '';
                if (nameConflicts.has(nameKey(data))) classes += ' conflict';
                if (dupLicUsers.has(data.Id)) classes += ' dup-license'; // Classe para linha toda
                if (classes) $(row).addClass(classes.trim());
                else $(row).removeClass('conflict dup-license');
            },
            drawCallback: function() {
                renderAlerts();
                hideLoader();
            }
        });
        $(table.table().node()).on('preDraw.dt', () => showLoader('Updating table...'));
    }

    // --- Lógica de Multi-Busca e Filtros da Tabela (mantida) ---
    function updateSearchFieldUI($row) { /* ... (mantido como na versão anterior) ... */ }
    function setupMultiSearch() { /* ... (mantido como na versão anterior) ... */ }
    function applyMultiSearch() { /* ... (mantido como na versão anterior) ... */ }
    function _executeMultiSearchLogic() { /* ... (mantido como na versão anterior) ... */ }
    $('#clearFilters').on('click', () => { /* ... (mantido como na versão anterior) ... */ });
    // --- Fim Lógica de Multi-Busca ---


    // --- Funções de Exportação (mantidas) ---
    function downloadCsv(csvContent, fileName) { /* ... (mantido) ... */ }
    $('#exportCsv').on('click', () => { /* ... (mantido) ... */ });
    $('#exportIssues').on('click', () => { /* ... (mantido, mas pode usar `dupLicUsers` para o relatório) ... */ });
    // --- Fim Funções de Exportação ---


    // --- Processamento de Dados Inicial com Web Worker (mantido) ---
    function processDataWithWorker(rawData) { /* ... (mantido como na versão anterior, uniqueFieldValues será usado pela IA) ... */ }
    // --- Fim Processamento de Dados Inicial ---


    // ===========================================================
    // FUNCIONALIDADE DE IA - RETROCONVERSAÇÃO E FILTRAGEM
    // ===========================================================
    const AI_API_KEY_STORAGE_KEY = 'licenseAiApiKey_v1_google_prompt_efficient';
    const $manageAiApiKeyButton = $('#manageAiApiKeyButton');
    const $askAiButton = $('#askAiButton');
    const $aiQuestionInput = $('#aiQuestionInput');
    const $aiResponseArea = $('#aiResponseArea');
    const $aiLoadingIndicator = $('#aiLoadingIndicator');
    const $aiTokenInfo = $('#aiTokenInfo');

    let aiConversationHistory = [];
    const MAX_CONVERSATION_HISTORY_TURNS = 2; // Reduzido para dar mais espaço ao contexto filtrado
    const TARGET_MAX_CHARS_FOR_API = 780000; // Alvo ~780k caracteres (Gemini 1.5 Flash tem 1M tokens ~ 3-4M chars)

    function getAiApiKey() { return localStorage.getItem(AI_API_KEY_STORAGE_KEY); }

    function updateAiApiKeyStatusDisplay() {
        // ... (mantido como na versão anterior, chama updateAiFeatureStatus) ...
        if (getAiApiKey()) {
            $manageAiApiKeyButton.text('API IA Config.');
            $manageAiApiKeyButton.css('background-color', '#28a745');
            $manageAiApiKeyButton.attr('title', 'Chave da API da IA (Google Gemini) configurada. Clique para alterar ou remover.');
        } else {
            $manageAiApiKeyButton.text('Configurar API IA');
            $manageAiApiKeyButton.css('background-color', ''); 
            $manageAiApiKeyButton.attr('title', 'Configurar chave da API da IA (Google Gemini) para análise.');
        }
        updateAiFeatureStatus(); 
    }

    function handleApiKeyInputViaPrompt() {
        // ... (mantido como na versão anterior, chama updateAiApiKeyStatusDisplay) ...
        const currentKey = getAiApiKey() || "";
        const newKey = window.prompt("Por favor, insira sua chave de API do Google Gemini:", currentKey);

        if (newKey === null) { 
        } else if (newKey.trim() === "" && currentKey !== "") {
            if (confirm("Você tem certeza que deseja remover a chave da API salva?")) {
                localStorage.removeItem(AI_API_KEY_STORAGE_KEY);
                alert("Chave da API removida.");
            }
        } else if (newKey.trim() !== "") {
            localStorage.setItem(AI_API_KEY_STORAGE_KEY, newKey.trim());
            alert("Chave da API da IA (Google Gemini) salva!");
        }
        updateAiApiKeyStatusDisplay();
        return getAiApiKey() !== null && getAiApiKey().trim() !== "";
    }

    $manageAiApiKeyButton.on('click', function() { handleApiKeyInputViaPrompt(); });

    function estimateChars(text) { return typeof text === 'string' ? text.length : 0; }

    async function updateAiFeatureStatus() {
        const apiKey = getAiApiKey();
        const hasData = allData && allData.length > 0;

        if (!apiKey || !hasData) {
            $askAiButton.prop('disabled', true);
            $askAiButton.attr('title', !apiKey ? 'Configure a chave da API primeiro.' : 'Carregue os dados primeiro.');
            $aiTokenInfo.text('');
            return;
        }

        // Simulação leve para o status do botão, a verificação completa ocorre ao perguntar
        const tempQuestionForEstimation = "Resumo geral."; // Pergunta genérica para estimativa
        const { currentChars, canSend } = await analyzeAndPrepareForAISending(tempQuestionForEstimation, false);


        if (!canSend) {
            $askAiButton.prop('disabled', true);
            $askAiButton.attr('title', `Os dados base excedem o limite (${Math.round(currentChars/1000)}k > ${Math.round(TARGET_MAX_CHARS_FOR_API/1000)}k chars). A IA pode não funcionar corretamente.`);
            $aiTokenInfo.text(`AVISO: Dados base (${Math.round(currentChars/1000)}k chars) podem exceder o limite para perguntas complexas.`).css('color', 'orange');
        } else {
            $askAiButton.prop('disabled', false);
            $askAiButton.attr('title', 'Perguntar à IA');
            $aiTokenInfo.text(`Pronto para IA. Estimativa base: ~${Math.round(currentChars/1000)}k / ${Math.round(TARGET_MAX_CHARS_FOR_API/1000)}k chars.`).css('color', '');
        }
    }

    // -------- INÍCIO: Lógica de Análise de Pergunta e Filtragem de Dados --------
    function findUsersWithDuplicateLicenses(sourceData, options = {}) {
        const { identifyOnly = false, paidOnly = false } = options;
        const usersWithDuplicates = [];
        const paidLicenseKeywords = ['E3', 'E5', 'Premium', 'P1', 'P2', 'Standard', 'Copilot']; // Definir melhor se houver campo específico

        sourceData.forEach(user => {
            const licenseCounts = {};
            let hasDuplicate = false;
            let duplicateDetails = {};

            (user.Licenses || []).forEach(license => {
                const licenseName = license.LicenseName;
                if (!licenseName) return;

                let isConsideredPaid = true; // Assume paid unless specified otherwise
                if (paidOnly) {
                    isConsideredPaid = paidLicenseKeywords.some(kw => licenseName.toLowerCase().includes(kw.toLowerCase())) &&
                                       !licenseName.toLowerCase().includes('free') &&
                                       !licenseName.toLowerCase().includes('trial');
                }

                if (isConsideredPaid) {
                    licenseCounts[licenseName] = (licenseCounts[licenseName] || 0) + 1;
                    if (licenseCounts[licenseName] > 1) {
                        hasDuplicate = true;
                        duplicateDetails[licenseName] = licenseCounts[licenseName];
                    }
                }
            });

            if (hasDuplicate) {
                if (identifyOnly) {
                    usersWithDuplicates.push({ Id: user.Id, DisplayName: user.DisplayName, OfficeLocation: user.OfficeLocation, Licenses: user.Licenses, duplicateLicenseDetails: duplicateDetails });
                } else {
                    // Para envio à IA, pode-se retornar uma cópia mais leve do usuário
                    usersWithDuplicates.push({
                        Id: user.Id,
                        DisplayName: user.DisplayName,
                        Email: user.Email,
                        JobTitle: user.JobTitle,
                        OfficeLocation: user.OfficeLocation,
                        Licenses: (user.Licenses || []).map(l => ({ LicenseName: l.LicenseName, SkuId: l.SkuId })), // Manter SkuId pode ser útil
                        duplicateLicenseDetails: duplicateDetails // Adiciona detalhes das duplicatas
                    });
                }
            }
        });
        return usersWithDuplicates;
    }

    function filterUsersByJobTitleKeyword(sourceData, keyword) {
        if (!keyword || typeof keyword !== 'string') return sourceData;
        const lowerKeyword = keyword.toLowerCase();
        return sourceData.filter(user => user.JobTitle && user.JobTitle.toLowerCase().includes(lowerKeyword));
    }
    
    function filterUsersByLicenseName(sourceData, licenseNameQuery) {
        if (!licenseNameQuery || typeof licenseNameQuery !== 'string') return sourceData;
        const lowerLicenseNameQuery = licenseNameQuery.toLowerCase();
        return sourceData.filter(user =>
            (user.Licenses || []).some(lic => lic.LicenseName && lic.LicenseName.toLowerCase().includes(lowerLicenseNameQuery))
        );
    }

    async function analyzeQueryAndFilterData(userQuestion) {
        const lowerQuestion = userQuestion.toLowerCase();
        let filteredData = null; // Inicia como null, se não houver filtro específico, não será usado diretamente no prompt como "dados filtrados"
        let analysisNotes = ""; // Notas sobre a análise para a IA

        // Exemplo 1: Licenças Duplicadas
        if (lowerQuestion.includes("licenç duplicad") || lowerQuestion.includes("usuário com mais de uma licença igual")) {
            analysisNotes += "A pergunta se refere a licenças duplicadas. ";
            const paidOnly = lowerQuestion.includes("paga") || lowerQuestion.includes("premium"); // simples heurística
            filteredData = findUsersWithDuplicateLicenses(allData, { paidOnly });
            analysisNotes += `Foram encontrados ${filteredData.length} usuários com licenças duplicadas ${paidOnly ? '(considerando apenas pagas)' : ''}. `;
            if (filteredData.length > 0) {
                 analysisNotes += `Exemplo de detalhe da primeira duplicata encontrada: ${JSON.stringify(filteredData[0].duplicateLicenseDetails)}. `;
            }
        }
        // Exemplo 2: Cargo Técnico
        else if (lowerQuestion.includes("cargo técnico") || (lowerQuestion.includes("quantos") && lowerQuestion.includes("técnico"))) {
            analysisNotes += "A pergunta se refere a usuários com cargo técnico. ";
            if (uniqueFieldValues.JobTitle && uniqueFieldValues.JobTitle.length > 0) {
                filteredData = filterUsersByJobTitleKeyword(allData, "técnico"); // Ou usar uma lista mais completa de palavras-chave
                analysisNotes += `Filtrando por cargos contendo 'técnico', foram encontrados ${filteredData.length} usuários. `;
            } else {
                analysisNotes += "Não há informações de 'JobTitle' suficientemente detalhadas ou disponíveis para filtrar por 'técnico' de forma confiável. ";
                // filteredData permanece null ou allData se a IA precisar dos dados gerais
            }
        }
        // Exemplo 3: Usuários com uma licença específica
        const licenseMentionMatch = lowerQuestion.match(/usuários com (a licença |licença )?['"]?([^'"?]+)['"]?/);
        if (!filteredData && licenseMentionMatch && licenseMentionMatch[2]) {
            const mentionedLicense = licenseMentionMatch[2].trim();
            analysisNotes += `A pergunta parece buscar usuários com a licença "${mentionedLicense}". `;
            filteredData = filterUsersByLicenseName(allData, mentionedLicense);
            analysisNotes += `Foram encontrados ${filteredData.length} usuários com essa licença (ou similar). `;
        }


        // Se nenhum filtro específico foi acionado, filteredData pode ser null.
        // A IA receberá essa nota e poderá usar estatísticas gerais se filteredData for null/vazio.
        if (filteredData && filteredData.length === 0 && analysisNotes.includes("Foram encontrados 0 usuários")) {
           analysisNotes += "Nenhum usuário correspondeu aos critérios de filtragem específicos. ";
        } else if (!filteredData && analysisNotes === "") {
            analysisNotes = "Nenhum filtro específico foi aplicado baseado na pergunta. A análise será geral ou baseada em palavras-chave diretas na pergunta. ";
        }


        return {
            originalQuestion: userQuestion,
            analyzedQueryNotes: analysisNotes.trim(), // Notas para a IA sobre o que foi filtrado
            dataForAI: filteredData // Pode ser null se nenhum filtro específico foi aplicado ou se o filtro não retornou resultados
        };
    }
    // -------- FIM: Lógica de Análise de Pergunta e Filtragem de Dados --------


    async function analyzeAndPrepareForAISending(userQuestion, isActualSendOperation) {
        let currentChars = estimateChars(JSON.stringify(aiConversationHistory)) + estimateChars(userQuestion);
        let contextForAI = "";
        let systemMessage = `Você é um assistente especialista em análise de licenciamento Microsoft 365.
Responda à pergunta do usuário baseando-se ESTRITAMENTE nos dados e notas de análise fornecidas.
Se os "Dados Filtrados Específicos" forem fornecidos, priorize-os.
Se os dados filtrados estiverem vazios ou não forem fornecidos, use as "Notas da Análise da Pergunta" para entender o contexto e, se necessário, informe que a filtragem não encontrou resultados ou que a pergunta requer uma análise geral das "Estatísticas Agregadas" (se disponíveis).
NÃO invente dados. Seja conciso e use Markdown. Não inclua saudações/despedidas.`;
        currentChars += estimateChars(systemMessage);

        // Passo 1: Análise da pergunta e filtragem de dados (CLIENT-SIDE)
        const { analyzedQueryNotes, dataForAI } = await analyzeQueryAndFilterData(userQuestion);

        contextForAI += `Notas da Análise da Pergunta (cliente): ${analyzedQueryNotes}\n---\n`;

        if (dataForAI !== null) { // Se dataForAI é null, significa que nenhum filtro específico foi aplicado com sucesso
            if (dataForAI.length > 0) {
                // Tentar enviar os dados filtrados, mas respeitando o limite de caracteres
                const jsonDataForAI = JSON.stringify(dataForAI.slice(0, 50)); // Envia uma amostra dos filtrados (até 50), ou todos se < 50
                const estimatedFilteredDataChars = estimateChars(jsonDataForAI);

                if (currentChars + estimatedFilteredDataChars < TARGET_MAX_CHARS_FOR_API) {
                    contextForAI += `Dados Filtrados Específicos (amostra de até 50 usuários correspondentes):\n${jsonDataForAI}\n---\n`;
                    currentChars += estimatedFilteredDataChars;
                } else {
                    contextForAI += `Dados Filtrados Específicos: A lista de ${dataForAI.length} usuários filtrados é muito grande para incluir diretamente. A IA deve usar as 'Notas da Análise' e o número de usuários encontrados para responder.\n---\n`;
                }
            } else { // dataForAI existe (ou seja, um filtro foi tentado) mas retornou vazio
                 contextForAI += "Dados Filtrados Específicos: Nenhum usuário encontrado após a filtragem aplicada pelo cliente.\n---\n";
            }
        }

        // Adicionar estatísticas agregadas gerais se houver espaço e nenhum dado filtrado útil
        if ((dataForAI === null || dataForAI.length === 0) && allData.length > 0) {
            const totalUsers = allData.length;
            const licenseMap = new Map();
            allData.forEach(user => (user.Licenses || []).forEach(lic => licenseMap.set(lic.LicenseName, (licenseMap.get(lic.LicenseName) || 0) + 1)));
            
            let aggregatedStats = `Estatísticas Agregadas Gerais (Total de ${totalUsers} usuários no dataset):\n`;
            aggregatedStats += `- Licenças Únicas Distribuídas: ${licenseMap.size}\n`;
            const topLicensesOverall = [...licenseMap.entries()].sort(([,a],[,b]) => b-a).slice(0,5).map(([n,c])=>`  - ${n}: ${c} usuários`).join('\n'); // Top 5
            aggregatedStats += `- Top 5 Licenças (Geral):\n${topLicensesOverall}\n`;
            // Adicionar mais estatísticas se necessário e couber
            
            const estimatedAggregatedChars = estimateChars(aggregatedStats);
            if (currentChars + estimatedAggregatedChars < TARGET_MAX_CHARS_FOR_API) {
                contextForAI += aggregatedStats;
                currentChars += estimatedAggregatedChars;
            } else {
                contextForAI += "Estatísticas Agregadas Gerais: Omitidas devido ao limite de tamanho do prompt.\n";
            }
        }
        
        const canSendData = currentChars < TARGET_MAX_CHARS_FOR_API;
        if (isActualSendOperation) {
            $aiTokenInfo.text(`Preparado: ~${Math.round(currentChars/1000)}k / ${Math.round(TARGET_MAX_CHARS_FOR_API/1000)}k chars. ${canSendData ? 'Pronto para envio.' : 'Excede o limite!'}`).css('color', canSendData ? 'green' : 'red');
        }
        
        return {
            contextForAI,
            systemMessage,
            history: aiConversationHistory,
            canSend: canSendData,
            currentChars // Retorna para o status do botão
        };
    }


    $askAiButton.on('click', async function() {
        let apiKey = getAiApiKey();
        if (!apiKey) { /* ... (lógica de pedir API key mantida) ... */ return; }

        const userQuestion = $aiQuestionInput.val().trim();
        if (!userQuestion) { /* ... (lógica de pergunta vazia mantida) ... */ return; }
        if (!allData || allData.length === 0) { /* ... (lógica de sem dados mantida) ... */ return; }

        $aiLoadingIndicator.removeClass('hidden').text('Analisando pergunta e preparando dados...');
        $askAiButton.prop('disabled', true);
        $aiResponseArea.html(`Analisando pergunta e preparando dados... <i class="fas fa-spinner fa-spin"></i>`);

        const { contextForAI, systemMessage, history, canSend } = await analyzeAndPrepareForAISending(userQuestion, true);

        if (!canSend) {
            $aiResponseArea.text('Os dados preparados para a IA excedem o limite de tamanho. Tente uma pergunta mais simples, filtre mais os dados na tabela ou limpe o histórico da IA (recarregando a página).');
            $aiLoadingIndicator.addClass('hidden');
            $askAiButton.prop('disabled', false);
            updateAiFeatureStatus();
            return;
        }
        
        $aiLoadingIndicator.text('Enviando para IA...');
        $aiResponseArea.html(`Enviando para IA (Google Gemini)... <i class="fas fa-spinner fa-spin"></i>`);

        const GEMINI_MODEL_TO_USE = "gemini-1.5-flash-latest";
        const AI_PROVIDER_ENDPOINT_BASE = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL_TO_USE}:generateContent`;
        const finalEndpoint = `${AI_PROVIDER_ENDPOINT_BASE}?key=${apiKey}`;

        try {
            const requestHeaders = { 'Content-Type': 'application/json', };
            let conversationTurnsForAPI = [];

            if (history.length === 0) { // Primeira pergunta da sessão
                conversationTurnsForAPI.push({ "role": "user", "parts": [{ "text": systemMessage }] });
                conversationTurnsForAPI.push({ "role": "model", "parts": [{ "text": "Entendido. Estou pronto para ajudar com a análise, usando os dados e notas fornecidas."}] });
            } else {
                conversationTurnsForAPI = [...history];
            }
            
            // O contextForAI já inclui as notas da análise e os dados filtrados (ou info sobre eles)
            const userTurn = { "role": "user", "parts": [{ "text": `${contextForAI}\n\nPERGUNTA ORIGINAL DO USUÁRIO: ${userQuestion}` }] };
            conversationTurnsForAPI.push(userTurn);
            
            const requestBody = {
              contents: conversationTurnsForAPI,
              generationConfig: { "maxOutputTokens": 3072, "temperature": 0.3 } // Aumentado um pouco maxOutputTokens
            };

            const response = await fetch(finalEndpoint, { /* ... (fetch e tratamento de erro mantidos) ... */ });
            const responseBodyText = await response.text(); // Mantido
            let data;  // Mantido
            try { data = JSON.parse(responseBodyText); } // Mantido
            catch (e) { /* ... (tratamento de erro de parse mantido) ... */ } // Mantido

            if (!response.ok) { /* ... (tratamento de erro da API mantido) ... */ } // Mantido

            let aiTextResponse = "Não foi possível extrair a resposta da IA ou a resposta estava vazia."; // Mantido
            if (data.candidates && data.candidates[0]) { /* ... (extração da resposta mantida) ... */ } // Mantido
            
            aiConversationHistory.push(userTurn);
            aiConversationHistory.push({ "role": "model", "parts": [{ "text": aiTextResponse }] });

            const maxHistoryItems = (MAX_CONVERSATION_HISTORY_TURNS * 2) + (aiConversationHistory.find(turn => turn.parts[0].text.includes("Você é um assistente especialista")) ? 2 : 0);
            if (aiConversationHistory.length > maxHistoryItems) {
                if (aiConversationHistory[0].parts[0].text.includes("Você é um assistente especialista")) {
                    const systemPromptAndFirstModelResponse = aiConversationHistory.slice(0, 2);
                    const recentTurns = aiConversationHistory.slice(-((MAX_CONVERSATION_HISTORY_TURNS -1) * 2));
                    aiConversationHistory = [...systemPromptAndFirstModelResponse, ...recentTurns];
                } else { 
                    aiConversationHistory = aiConversationHistory.slice(-(MAX_CONVERSATION_HISTORY_TURNS * 2));
                }
            }
            
            $aiQuestionInput.val('');
            function simpleMarkdownToHtml(mdText) { /* ... (mantido) ... */ } // Mantido
            $aiResponseArea.html(simpleMarkdownToHtml(aiTextResponse));

        } catch (error) { /* ... (tratamento de erro mantido) ... */ }
        finally {
            $aiLoadingIndicator.addClass('hidden');
            $askAiButton.prop('disabled', false);
            updateAiFeatureStatus();
        }
    });
   
    // --- Inicialização ---
    try {
        if (typeof userData !== 'undefined' && ( (Array.isArray(userData) && userData.length > 0) || (userData.data && Array.isArray(userData.data) && userData.data.length > 0) ) ) {
            // ... (lógica de carregamento inicial mantida, chama processDataWithWorker)
            let dataToProcess = Array.isArray(userData) ? userData : userData.data;
            if (userData.error || (Array.isArray(dataToProcess) && dataToProcess.length === 0 && !userData.message) ) {
                 $('#searchCriteria').text(userData.error || 'JSON data is empty or invalid as provided by PowerShell.');
                 console.error("Error in userData from PowerShell or empty data:", userData);
                 hideLoader();
                 updateAiFeatureStatus();
            } else if (userData.message && Array.isArray(dataToProcess) && dataToProcess.length === 0) {
                 $('#searchCriteria').text(userData.message);
                 hideLoader();
                 updateAiFeatureStatus();
            }
            else {
                processDataWithWorker(dataToProcess); // Isso vai popular allData e uniqueFieldValues
            }
        } else if (typeof userData !== 'undefined' && userData.message) {
             $('#searchCriteria').text(userData.message);
             hideLoader();
             updateAiFeatureStatus();
        } else { /* ... (tratamento de userData não definido mantido) ... */ }
    } catch (error) { /* ... (tratamento de erro inicial mantido) ... */ }
    updateAiApiKeyStatusDisplay(); // Chamada inicial
    // ===========================================================
    // FIM FUNCIONALIDADE DE IA
    // ===========================================================

}); // Fim do $(document).ready()
