document.addEventListener("DOMContentLoaded", function() {

  fetch('http://localhost:5000/test-results')
    .then(response => response.json())
    .then(data => {
      renderTable(data);
    })
    .catch(error => console.error('Erro ao carregar o JSON:', error));

  function renderTable(data) {
    const tableBody = document.getElementById('test-results-body');
    tableBody.innerHTML = ''; // Limpar tabela antes de preencher

    // Ir buscar user
    const user_file = data.config.configFile;

    for (var i = 0; i < user_file.length; i++)
    {
      
    }

    // Iterar sobre cada suite no JSON
    data.suites.forEach(suite => {
      // Iterar sobre cada spec dentro da suite
      suite.specs.forEach(spec => {
        
        // Criar uma nova linha na tabela
        const row = document.createElement('tr');

        //Ir buscar Titulo do teste e da página correspondente
        const tituloColumn = document.createElement('td');
        tituloColumn.innerHTML = `<h6><b>${spec.title}</b></h6><br>${suite.title}`;
        row.appendChild(tituloColumn);

        //Ir buscar Duração
        const duracaoColumn = document.createElement('td');
        const duracao = spec.tests[0].results[0].duration;
        const segundos = (duracao / 1000).toFixed(2);
        duracaoColumn.innerHTML = segundos + ' s';
        row.appendChild(duracaoColumn);

        //Ir buscar Estado
        const estadoColumn = document.createElement('td');
        estadoColumn.innerHTML = spec.tests[0].results[0].status;
        row.appendChild(estadoColumn);

        //Ir buscar Erro
        const erroColumn = document.createElement('td');
        const verificar_erro = spec.tests[0].results[0].error;
        if (verificar_erro != null) erroColumn.innerHTML = spec.tests[0].results[0].error.message;
        row.appendChild(erroColumn);

        // Adicionar a linha à tabela
        tableBody.appendChild(row);
      });
    });
  }
});

document.getElementById("meuElemento").onclick = function() {
  //Por acabar
};