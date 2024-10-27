// Flag para controle de visibilidade da seção principal "HAOC"
let showHAOCSection = true;

// Função para alternar visibilidade da seção "HAOC"
function toggleHAOCSection() {
    showHAOCSection = !showHAOCSection;
    const haocContainer = document.getElementById('organogram-container');
    haocContainer.style.display = showHAOCSection ? 'block' : 'none';
    document.getElementById('toggleHAOCButton').textContent = showHAOCSection ? 'Ocultar HAOC' : 'Mostrar HAOC';
}

// Botão para alternar visibilidade da seção HAOC
const toggleButton = document.createElement('button');
toggleButton.id = 'toggleHAOCButton';
toggleButton.textContent = 'Ocultar HAOC';
toggleButton.onclick = toggleHAOCSection;
document.body.insertBefore(toggleButton, document.body.firstChild);

// Função para carregar o Excel e gerar o organograma
document.getElementById('loadExcel').addEventListener('click', function () {
    var input = document.getElementById('fileUpload');
    var reader = new FileReader();

    reader.onload = function (e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, { type: 'binary' });
        var sheetName = workbook.SheetNames[0];
        var sheet = workbook.Sheets[sheetName];
        var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Verificar e ajustar a estrutura do Excel para garantir compatibilidade
        const headers = jsonData[0].map(h => h.trim().toLowerCase());
        
        // Mapeia os índices de cada coluna necessária
        const nameIndex = headers.indexOf("nome do funcionário");
        const managerIndex = headers.indexOf("chefe imediato");
        const roleIndex = headers.indexOf("função");
        const departmentIndex = headers.indexOf("setor");
        const directorateIndex = headers.indexOf("diretoria");
        const matriculaIndex = headers.indexOf("matrícula");

        if (nameIndex === -1 || managerIndex === -1 || roleIndex === -1 || directorateIndex === -1 || matriculaIndex === -1) {
            alert("O arquivo não está no formato correto. Certifique-se de que o cabeçalho contém as colunas 'Nome do Funcionário', 'Chefe Imediato', 'Função', 'Setor', 'Diretoria', e 'Matrícula'.");
            return;
        }

        // Processar o jsonData para gerar o organograma
        generateOrganogram(jsonData, { nameIndex, managerIndex, roleIndex, departmentIndex, directorateIndex, matriculaIndex });
    };

    if (input.files.length) {
        reader.readAsBinaryString(input.files[0]);
    } else {
        alert('Selecione um arquivo para carregar.');
    }
});

// Função para remover acentos e caracteres especiais
function removeAccents(str) {
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// Função para capitalizar a primeira letra de cada palavra, evitando números romanos de I a V
function capitalizeWords(str) {
    const romanNumerals = ['I', 'II', 'III', 'IV', 'V'];
    return str
        .toLowerCase()
        .split(' ')
        .map(word => romanNumerals.includes(word.toUpperCase()) ? word.toUpperCase() : word.charAt(0).toUpperCase() + word.slice(1))
        .join(' ');
}

// Função para abreviar o nome da diretoria
function abbreviateDirectorate(name) {
    return name.split(' ').map(word => word.charAt(0)).join('').toUpperCase();
}

// Variáveis globais para organizar dados de hierarquia
let hierarchy = {};
let roles = {};
let directorates = {};
let matriculas = {}; // Novo objeto para armazenar matrículas

// Função para contar subordinados de um gerente
function countSubordinates(manager) {
    if (!hierarchy[manager]) return 0;
    return hierarchy[manager].reduce((count, subordinate) => {
        return count + 1 + countSubordinates(subordinate.name);
    }, 0);
}

// Função principal para gerar o organograma com classificação alfabética
function generateOrganogram(data, indices) {
    let container = document.getElementById('organogram-container');
    container.innerHTML = ''; // Limpa o container

    hierarchy = {};
    roles = {};
    directorates = {};
    matriculas = {}; // Limpar a lista de matrículas

    // Organizar os dados por chefe e consolidar subordinados
    data.slice(1).forEach(row => {
        let employee = removeAccents(row[indices.nameIndex]);
        let manager = removeAccents(row[indices.managerIndex]);
        let role = removeAccents(row[indices.roleIndex]);
        let directorate = row[indices.directorateIndex];
        let matricula = row[indices.matriculaIndex];

        if (!roles[employee]) roles[employee] = role;
        if (!directorates[employee]) directorates[employee] = directorate;
        if (!matriculas[employee]) matriculas[employee] = matricula;

        if (!hierarchy[manager]) {
            hierarchy[manager] = [];
        }
        hierarchy[manager].push({
            name: employee,
            role: role,
            directorate: directorate,
            matricula: matricula
        });
    });

    // Classifica os gestores de nível superior (sem gerentes acima)
    let topLevelManagers = Object.keys(hierarchy).filter(manager => !data.slice(1).some(row => removeAccents(row[indices.nameIndex]) === manager));
    topLevelManagers.sort((a, b) => a.localeCompare(b));

    // Criar a linha inicial com todos os gestores de nível superior
    let topLevelContainer = document.createElement('div');
    topLevelContainer.classList.add('top-level-container');

    topLevelManagers.forEach(manager => {
        let subordinatesCount = countSubordinates(manager); // Calcula o número de subordinados
        let managerDiv = createManagerNode(manager, roles[manager], directorates[manager], subordinatesCount, matriculas[manager]);
        topLevelContainer.appendChild(managerDiv);

        // Adicionar container para subordinados (inicialmente fechado)
        let subordinatesContainer = document.createElement('div');
        subordinatesContainer.classList.add('subordinates-container');
        subordinatesContainer.style.display = 'none';

        if (hierarchy[manager]) {
            hierarchy[manager].sort((a, b) => a.name.localeCompare(b.name));
            hierarchy[manager].forEach(subordinate => {
                let subordinateTree = createSubordinateTree(subordinate.name, hierarchy, roles, directorates, matriculas);
                subordinatesContainer.appendChild(subordinateTree);
            });
        }

        managerDiv.appendChild(subordinatesContainer);
    });

    container.appendChild(topLevelContainer);
}

// Função para criar o nó de um gestor com a contagem de subordinados e a matrícula
function createManagerNode(manager, role, directorate, subordinatesCount, matricula) {
    let managerDiv = document.createElement('div');
    managerDiv.classList.add('node', 'manager');

    let directorateAbbr = abbreviateDirectorate(capitalizeWords(directorate || ''));

    managerDiv.innerHTML = `
        <div class="employee-name">
            <strong>${capitalizeWords(manager)}</strong><br>
            <span class="role">${capitalizeWords(role || '')}</span><br>
            <span class="directorate" title="${directorate}">${directorateAbbr}</span><br>
            <span class="matricula"><strong>Matrícula: ${matricula || '-'}</strong></span><br>
            <span class="headcount-square">${subordinatesCount === 0 ? '-' : subordinatesCount}</span>
        </div>
        <button class="toggle-btn">+</button>
    `;

    managerDiv.querySelector('.toggle-btn').addEventListener('click', function () {
        const subordinatesContainer = managerDiv.querySelector('.subordinates-container');
        const isVisible = subordinatesContainer.style.display === 'flex';
        subordinatesContainer.style.display = isVisible ? 'none' : 'flex';
        this.textContent = isVisible ? '+' : '-';

        if (!isVisible) {
            managerDiv.classList.add('highlight');
        } else {
            managerDiv.classList.remove('highlight');
        }
    });

    return managerDiv;
}

// Função para criar a árvore de subordinados com classificação e matrícula
function createSubordinateTree(name, hierarchy, roles, directorates, matriculas) {
    let subordinateDiv = document.createElement('div');
    subordinateDiv.classList.add('node', 'employee');

    let subordinatesCount = countSubordinates(name); // Contagem de subordinados
    let directorateAbbr = abbreviateDirectorate(capitalizeWords(directorates[name] || ''));

    subordinateDiv.innerHTML = `
        <div class="employee-name">
            <strong>${capitalizeWords(name)}</strong><br>
            <span class="role">${capitalizeWords(roles[name] || '')}</span><br>
            <span class="directorate" title="${directorates[name] || ''}">${directorateAbbr}</span><br>
            <span class="matricula"><strong>Matrícula: ${matriculas[name] || '-'}</strong></span><br>
            <span class="headcount-square">${subordinatesCount === 0 ? '-' : subordinatesCount}</span>
        </div>
    `;

    if (hierarchy[name]) {
        let subordinatesContainer = document.createElement('div');
        subordinatesContainer.classList.add('subordinates-container');
        subordinatesContainer.style.display = 'none';

        // Classifica os subordinados em ordem alfabética
        hierarchy[name].sort((a, b) => a.name.localeCompare(b.name));
        hierarchy[name].forEach(subordinate => {
            let childSubordinate = createSubordinateTree(subordinate.name, hierarchy, roles, directorates, matriculas);
            subordinatesContainer.appendChild(childSubordinate);
        });

        subordinateDiv.appendChild(subordinatesContainer);

        let toggleButton = document.createElement('button');
        toggleButton.classList.add('toggle-btn');
        toggleButton.textContent = '+';
        toggleButton.addEventListener('click', function () {
            const isVisible = subordinatesContainer.style.display === 'flex';
            subordinatesContainer.style.display = isVisible ? 'none' : 'flex';
            this.textContent = isVisible ? '+' : '-';

            if (!isVisible) {
                subordinateDiv.classList.add('highlight');
            } else {
                subordinateDiv.classList.remove('highlight');
            }
        });

        subordinateDiv.appendChild(toggleButton);
    }

    return subordinateDiv;
}

// Funções para expandir e colapsar todos os nós
function expandAll() {
    document.querySelectorAll('.subordinates-container').forEach(container => {
        container.style.display = 'flex';
    });
    document.querySelectorAll('.toggle-btn').forEach(button => {
        button.textContent = '-';
    });
}

function collapseAll() {
    document.querySelectorAll('.subordinates-container').forEach(container => {
        container.style.display = 'none';
    });
    document.querySelectorAll('.toggle-btn').forEach(button => {
        button.textContent = '+';
    });
}

// Função para exportar o organograma para PDF
function exportToPDF() {
    const element = document.getElementById('organogram-container');
    const options = {
        margin:       0.5,
        filename:     'organogram.pdf',
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2 },
        jsPDF:        { unit: 'in', format: 'a4', orientation: 'portrait' }
    };
    html2pdf().set(options).from(element).save();
}

// Adicionar botão para exportação para PDF
const exportButton = document.createElement('button');
exportButton.id = 'exportToPDFButton';
exportButton.textContent = 'Exportar para PDF';
exportButton.onclick = exportToPDF;
document.body.insertBefore(exportButton, document.body.firstChild);

// Função de pesquisa para localizar um funcionário ou chefe
function searchOrganogram() {
    let query = removeAccents(document.getElementById('searchInput').value.trim().toLowerCase());
    let found = false;

    document.querySelectorAll('.highlight-search').forEach(node => {
        node.classList.remove('highlight-search');
    });

    document.querySelectorAll('.employee-name').forEach(node => {
        let name = removeAccents(node.textContent.trim().toLowerCase());
        if (name.includes(query)) {
            expandToNode(node);
            node.scrollIntoView({ behavior: 'smooth', block: 'center' });
            node.classList.add('highlight-search');
            found = true;
        }
    });

    if (!found) {
        alert('Nenhum resultado encontrado');
    }
}

// Função auxiliar para expandir até o nó encontrado
function expandToNode(node) {
    let parent = node.closest('.subordinates-container');
    while (parent) {
        parent.style.display = 'block';
        let toggleButton = parent.previousElementSibling.querySelector('.toggle-btn');
        if (toggleButton) toggleButton.textContent = '-';
        parent = parent.closest('.subordinates-container').parentNode.closest('.subordinates-container');
    }
}


// Função para gerar e baixar o template de Excel
function downloadTemplate() {
    // Dados do template com nomes genéricos
    const templateData = [
        ["Nome do Funcionário", "Chefe Imediato", "Função", "Setor", "Diretoria", "Matrícula"],
        ["Funcionário A", "Funcionário B", "Analista", "Setor X", "Diretoria Y", "10001"],
        ["Funcionário B", "", "Gerente", "Setor X", "Diretoria Y", "10002"],
        ["Funcionário C", "Funcionário B", "Assistente", "Setor Z", "Diretoria Y", "10003"]
        ["..."],
        ["Instrução: Certifique-se de que o nome do 'Chefe Imediato' seja exatamente igual ao 'Nome do Funcionário' que ocupa essa posição. Qualquer diferença, mesmo de espaços ou letras maiúsculas/minúsculas, pode causar problemas na hierarquia."],
        
    ];

    // Criação do workbook e da planilha
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(templateData);

    // Adiciona a planilha ao workbook
    XLSX.utils.book_append_sheet(wb, ws, "Template Organograma");

    // Gera o arquivo para download
    XLSX.writeFile(wb, "Template_Organograma.xlsx");
}

// Adiciona o evento de clique para download do template
document.getElementById('downloadTemplate').addEventListener('click', downloadTemplate);


// Evento de busca ao pressionar "Enter"
document.getElementById('searchInput').addEventListener('keydown', function (event) {
    if (event.key === 'Enter') {
        searchOrganogram();
    }
});
