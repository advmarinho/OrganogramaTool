// Flag para controle de visibilidade da seção principal "HAOC"
let showHAOCSection = false; // Define se a seção "HAOC" será exibida por padrão

// Função para alternar visibilidade da seção "HAOC"
function toggleHAOCSection() {
    showHAOCSection = !showHAOCSection; // Inverte o valor da flag
    const haocContainer = document.getElementById('haoc-container');
    haocContainer.style.display = showHAOCSection ? 'block' : 'none'; // Mostra ou oculta com base na flag
    document.getElementById('toggleHAOCButton').textContent = showHAOCSection ? 'Ocultar HAOC' : 'Mostrar HAOC';
}

// Botão para alternar visibilidade da seção HAOC
const toggleButton = document.createElement('button');
toggleButton.id = 'toggleHAOCButton';
toggleButton.textContent = 'Ocultar HAOC';
toggleButton.onclick = toggleHAOCSection;
document.body.appendChild(toggleButton);

document.getElementById('loadExcel').addEventListener('click', function () {
    var input = document.getElementById('fileUpload');
    var reader = new FileReader();

    reader.onload = function (e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, { type: 'binary' });
        var sheetName = workbook.SheetNames[0];
        var sheet = workbook.Sheets[sheetName];
        var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Processar o jsonData para gerar o organograma
        generateOrganogram(jsonData);
        displayDirectorates(jsonData); // Adiciona a exibição das diretorias separadas
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

// Função para capitalizar a primeira letra de cada palavra,
// mantendo números romanos e preposições inalterados
function capitalizeWords(str) {
    const romanNumerals = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"];
    return str.toLowerCase().replace(/\b([a-zçáéíóú]+|[ivx]+)\b/g, (word) => {
        if (romanNumerals.includes(word)) {
            return word.toUpperCase();
        }
        if (["de", "do", "da", "e"].includes(word)) {
            return word;
        }
        return word.charAt(0).toUpperCase() + word.slice(1);
    });
}

function generateOrganogram(data) {
    let container = document.createElement('div');
    container.id = 'haoc-container'; // Contêiner para a seção "HAOC"
    container.style.display = showHAOCSection ? 'block' : 'none'; // Aplica a visibilidade com base na flag
    document.body.appendChild(container);
    container.innerHTML = ''; // Limpa o container

    let hierarchy = {};
    let totalCount = 0; // Contador geral de todos os funcionários e chefias

    // Organizar os dados por chefe e consolidar subordinados
    data.slice(1).forEach(row => {
        let employee = removeAccents(row[0]);
        let manager = removeAccents(row[1]);
        let role = removeAccents(row[2]);
        let department = removeAccents(row[3]);
        let directorate = removeAccents(row[4]);

        if (!hierarchy[manager]) {
            hierarchy[manager] = [];
        }

        hierarchy[manager].push({
            name: employee,
            role: role,
            department: department,
            directorate: directorate
        });
        totalCount++; // Incrementa o contador geral
    });

    // Função recursiva para criar a hierarquia de cada chefe, com contagem de subordinados
    function createHierarchy(manager, employees) {
        employees.sort((a, b) => a.name.localeCompare(b.name)); // Ordena chefias subordinadas

        let managerDiv = document.createElement('div');
        let managerCount = manager === 'HAOC' ? totalCount : employees.length;

        managerDiv.classList.add('node', 'manager');
        managerDiv.innerHTML = `
            <strong class="employee-name">${capitalizeWords(manager)}</strong>
            <span class="counter">${managerCount > 0 ? managerCount : '-'}</span>
            <button class="toggle-btn">+</button>
        `;

        let subordinatesContainer = document.createElement('div');
        subordinatesContainer.classList.add('subordinates-container');
        subordinatesContainer.style.display = 'none';

        employees.forEach(subordinate => {
            let subordinateDiv = document.createElement('div');
            subordinateDiv.classList.add('node', 'employee');

            if (hierarchy[subordinate.name]) {
                subordinateDiv.appendChild(createHierarchy(subordinate.name, hierarchy[subordinate.name]));
            } else {
                let displayName = capitalizeWords(subordinate.name);
                if (displayName.includes("Oswaldo Cruz")) {
                    displayName = displayName.replace("Oswaldo Cruz", "<span style='color: green; font-weight: bold;'>Oswaldo Cruz</span>");
                }
                
                subordinateDiv.innerHTML = `
                    <p class="employee-name"><strong>Nome:</strong> ${displayName} <br>
                    <strong>Função:</strong> ${capitalizeWords(subordinate.role)} <br>
                    <strong>Setor:</strong> ${capitalizeWords(subordinate.department)} <br>
                    <strong>Diretoria:</strong> ${capitalizeWords(subordinate.directorate)}</p>
                `;
            }

            subordinatesContainer.appendChild(subordinateDiv);
        });

        managerDiv.appendChild(subordinatesContainer);

        managerDiv.querySelector('.toggle-btn').addEventListener('click', function () {
            const isVisible = subordinatesContainer.style.display === 'block';
            subordinatesContainer.style.display = isVisible ? 'none' : 'block';
            this.textContent = isVisible ? '+' : '-';
        });

        return managerDiv;
    }

    let sortedManagers = Object.keys(hierarchy).sort((a, b) => {
        if (a === "HAOC") return -1;
        if (b === "HAOC") return 1;
        return a.localeCompare(b); // Ordena alfabeticamente os gestores
    });

    sortedManagers.forEach(manager => {
        let hierarchyTree = createHierarchy(manager, hierarchy[manager]);
        container.appendChild(hierarchyTree);
    });
}

// Função para exibir as diretorias separadas, com "Jose Marcelo Amatuzzi de Oliveira" no topo e até três níveis de subordinados
function displayDirectorates(data) {
    let directorateContainer = document.createElement('div');
    directorateContainer.id = 'directorate-container';
    document.body.appendChild(directorateContainer);

    let directorates = {};

    // Organizar dados por diretoria
    data.slice(1).forEach(row => {
        let directorate = capitalizeWords(removeAccents(row[4])); // Diretoria na coluna 4
        let employee = capitalizeWords(removeAccents(row[0]));
        let manager = capitalizeWords(removeAccents(row[1]));
        let role = capitalizeWords(removeAccents(row[2]));
        let department = capitalizeWords(removeAccents(row[3]));

        // Inicializar diretoria se não existir
        if (!directorates[directorate]) {
            directorates[directorate] = [];
        }

        directorates[directorate].push({
            name: employee,
            manager: manager,
            role: role,
            department: department,
            directorate: directorate
        });
    });

    // Exibir cada diretoria e seus funcionários, com "Jose Marcelo Amatuzzi de Oliveira" no topo
    Object.keys(directorates).sort().forEach(directorate => {
        let directorateDiv = document.createElement('div');
        directorateDiv.classList.add('directorate-section');
        directorateDiv.innerHTML = `<h2>${directorate}</h2>`;

        // Primeiro, exibe os diretores que reportam a "Jose Marcelo Amatuzzi de Oliveira"
        let directReports = directorates[directorate].filter(person => person.manager === "Jose Marcelo Amatuzzi de Oliveira");
        directReports.sort((a, b) => a.name.localeCompare(b.name)); // Ordena os diretores
        directReports.forEach(director => {
            let directorDiv = createHierarchySection(director, directorates[directorate]);
            directorateDiv.appendChild(directorDiv);
        });

        directorateContainer.appendChild(directorateDiv);
    });
}

// Função para criar uma seção de hierarquia recursiva com até três níveis de subordinados
function createHierarchySection(person, directorateData) {
    let personDiv = document.createElement('div');
    personDiv.classList.add('node', 'manager');
    personDiv.innerHTML = `
        <strong class="employee-name">Nome: ${person.name}</strong><br>
        <strong>Função:</strong> ${person.role} <br>
        <strong>Setor:</strong> ${person.department} <br>
        <strong>Diretoria:</strong> ${person.directorate} <br>
        <span class="counter">${countSubordinates(person.name, directorateData) > 0 ? countSubordinates(person.name, directorateData) : '-'}</span>
        <button class="toggle-btn">+</button>
    `;

    let subordinatesContainer = document.createElement('div');
    subordinatesContainer.classList.add('subordinates-container');
    subordinatesContainer.style.display = 'none';

    // Adiciona subordinados do primeiro nível
    let firstLevelSubordinates = directorateData.filter(sub => sub.manager === person.name);
    firstLevelSubordinates.sort((a, b) => a.name.localeCompare(b.name)); // Ordena os subordinados
    firstLevelSubordinates.forEach(firstSub => {
        let firstSubDiv = createSubordinateSection(firstSub, directorateData, 2); // Passa para o segundo nível
        subordinatesContainer.appendChild(firstSubDiv);
    });

    personDiv.appendChild(subordinatesContainer);

    personDiv.querySelector('.toggle-btn').addEventListener('click', function () {
        const isVisible = subordinatesContainer.style.display === 'block';
        subordinatesContainer.style.display = isVisible ? 'none' : 'block';
        this.textContent = isVisible ? '+' : '-';
    });

    return personDiv;
}

// Função para criar uma seção para subordinados recursivamente até três níveis
function createSubordinateSection(person, directorateData, level) {
    let personDiv = document.createElement('div');
    personDiv.classList.add('node', 'employee');
    personDiv.innerHTML = `
        <p class="employee-name"><strong>Nome:</strong> ${person.name} <br>
        <strong>Função:</strong> ${person.role} <br>
        <strong>Setor:</strong> ${person.department} <br>
        <strong>Diretoria:</strong> ${person.directorate} <br>
        <span class="counter">${countSubordinates(person.name, directorateData) > 0 ? countSubordinates(person.name, directorateData) : '-'}</span></p>
    `;

    // Adiciona subordinados do próximo nível, se houver
    if (level < 4) { // Limita ao terceiro nível
        let subordinatesContainer = document.createElement('div');
        subordinatesContainer.classList.add('subordinates-container');
        subordinatesContainer.style.display = 'none';

        let nextLevelSubordinates = directorateData.filter(sub => sub.manager === person.name);
        nextLevelSubordinates.sort((a, b) => a.name.localeCompare(b.name)); // Ordena os subordinados
        nextLevelSubordinates.forEach(nextSub => {
            let nextSubDiv = createSubordinateSection(nextSub, directorateData, level + 1);
            subordinatesContainer.appendChild(nextSubDiv);
        });

        if (nextLevelSubordinates.length > 0) {
            let toggleBtn = document.createElement('button');
            toggleBtn.classList.add('toggle-btn');
            toggleBtn.textContent = '+';
            toggleBtn.addEventListener('click', function () {
                const isVisible = subordinatesContainer.style.display === 'block';
                subordinatesContainer.style.display = isVisible ? 'none' : 'block';
                this.textContent = isVisible ? '+' : '-';
            });
            personDiv.appendChild(toggleBtn);
            personDiv.appendChild(subordinatesContainer);
        }
    }

    return personDiv;
}

// Função para contar subordinados de um determinado gerente
function countSubordinates(managerName, directorateData) {
    return directorateData.filter(sub => sub.manager === managerName).length;
}

// Funções para abrir e fechar todos os nós
function expandAll() {
    document.querySelectorAll('.subordinates-container').forEach(container => {
        container.style.display = 'block';
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

// Função de pesquisa para localizar um funcionário ou chefe
function searchOrganogram() {
    let query = removeAccents(document.getElementById('searchInput').value.trim().toLowerCase());
    let found = false;

    document.querySelectorAll('.highlight').forEach(node => {
        node.classList.remove('highlight');
    });

    document.querySelectorAll('.employee-name').forEach(node => {
        let name = removeAccents(node.textContent.trim().toLowerCase());
        if (name.includes(query)) {
            expandToNode(node);
            node.scrollIntoView({ behavior: 'smooth', block: 'center' });
            node.classList.add('highlight');
            found = true;
        }
    });

    if (!found) {
        alert('Nenhum resultado encontrado');
    }
}

function expandToNode(node) {
    let parent = node.closest('.subordinates-container');
    while (parent) {
        parent.style.display = 'block';
        let toggleButton = parent.previousElementSibling.querySelector('.toggle-btn');
        if (toggleButton) toggleButton.textContent = '-';
        parent = parent.closest('.subordinates-container').parentNode.closest('.subordinates-container');
    }
}

document.getElementById('searchInput').addEventListener('keydown', function(event) {
    if (event.key === 'Enter') {
        searchOrganogram();
    }
});

const style = document.createElement('style');
style.innerHTML = `
    .counter {
        background-color: #f0f0f0;
        color: #333;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: bold;
        display: inline-block;
        margin-left: 8px;
    }
    .highlight {
        background-color: #33cccc;
        font-weight: bold;
        border-radius: 4px;
        padding: 2px;
    }
    .directorate-section {
        margin-top: 20px;
        padding: 10px;
        border: 1px solid #ccc;
        border-radius: 8px;
        background-color: #f9f9f9;
    }
    .employee-entry {
        margin-bottom: 10px;
    }
`;
document.head.appendChild(style);
