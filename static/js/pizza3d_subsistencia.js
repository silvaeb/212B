// Inclui e inicializa o gráfico pizza 3D inclinado e vitrificado usando Plotly.js
// Os dados são passados via template Jinja2
function renderPizza3DSubsistencia(data) {
    // Cores sólidas e gradientes para simular 3D
    var colors = [
        'url(#gradYellow)',
        'url(#gradCyan)',
        'url(#gradRed)',
        'url(#gradBlue)',
        'url(#gradGreen)',
        'url(#gradOrange)'
    ];
    var pulls = [0.13, 0.11, 0.14, 0.12, 0.13, 0.11];
    var borderColors = [
        'rgba(60,60,60,0.22)',
        'rgba(60,60,60,0.22)',
        'rgba(60,60,60,0.22)',
        'rgba(60,60,60,0.22)',
        'rgba(60,60,60,0.22)',
        'rgba(60,60,60,0.22)'
    ];
    var trace = {
        type: 'pie',
        labels: data.labels,
        values: data.values,
        textinfo: 'label+percent',
        textposition: 'inside',
        marker: {
            colors: [
                'rgba(255,214,0,1)',
                'rgba(0,191,174,1)',
                'rgba(255,82,82,1)',
                'rgba(83,109,254,1)',
                'rgba(0,200,83,1)',
                'rgba(255,109,0,1)'
            ],
            line: {color: borderColors, width: 8},
            opacity: 1
        },
        hole: 0,
        pull: pulls,
        direction: 'clockwise',
        rotation: 300, // inclinação visual
        sort: false,
        showlegend: false,
        automargin: true
    };
    var layout = {
        title: {
            text: 'Distribuição Percentual dos Recursos',
            font: {size: 22, color: '#fff'},
            x: 0.5
        },
        paper_bgcolor: 'rgba(0,0,0,0)',
        plot_bgcolor: 'rgba(0,0,0,0)',
        showlegend: false,
        margin: {t: 60, l: 0, r: 0, b: 0},
        annotations: [
            {
                text: 'Valores em R$',
                font: {size: 13, color: '#FFD600', family: 'Montserrat, Arial'},
                showarrow: false,
                x: 0.5,
                y: -0.13,
                xref: 'paper',
                yref: 'paper'
            }
        ],
        // Sombra "abaixo" do gráfico
        shapes: [
            {
                type: 'ellipse',
                xref: 'paper', yref: 'paper',
                x0: 0.18, y0: 0.08, x1: 0.82, y1: 0.18,
                fillcolor: 'rgba(60,60,60,0.22)',
                line: {width: 0},
                layer: 'below'
            }
        ]
    };
    Plotly.newPlot('pizza3d-subsistencia', [trace], layout, {displayModeBar: false, responsive: true});
    // Efeito "volume" e brilho fake
    var el = document.getElementById('pizza3d-subsistencia');
    if (el) {
        el.style.boxShadow = '0 18px 38px 0 rgba(31,38,135,0.18), 0 2.5px 18px 0 rgba(60,60,60,0.13)';
        el.style.background = 'transparent';
        el.style.borderRadius = '30px';
        el.style.transition = 'box-shadow 0.7s cubic-bezier(.4,2,.6,1), background 0.7s';
    }
    // Adiciona gradientes SVG para simular 3D
    setTimeout(function() {
        var svg = el.querySelector('svg');
        if (svg && !svg.querySelector('#gradYellow')) {
            var defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
            defs.innerHTML = `
                <radialGradient id="gradYellow" cx="50%" cy="60%" r="80%">
                    <stop offset="0%" stop-color="#fff" stop-opacity="0.85"/>
                    <stop offset="60%" stop-color="#FFD600" stop-opacity="0.95"/>
                    <stop offset="100%" stop-color="#B89B00" stop-opacity="1"/>
                </radialGradient>
                <radialGradient id="gradCyan" cx="50%" cy="60%" r="80%">
                    <stop offset="0%" stop-color="#fff" stop-opacity="0.85"/>
                    <stop offset="60%" stop-color="#00BFAE" stop-opacity="0.95"/>
                    <stop offset="100%" stop-color="#00796B" stop-opacity="1"/>
                </radialGradient>
                <radialGradient id="gradRed" cx="50%" cy="60%" r="80%">
                    <stop offset="0%" stop-color="#fff" stop-opacity="0.85"/>
                    <stop offset="60%" stop-color="#FF5252" stop-opacity="0.95"/>
                    <stop offset="100%" stop-color="#B71C1C" stop-opacity="1"/>
                </radialGradient>
                <radialGradient id="gradBlue" cx="50%" cy="60%" r="80%">
                    <stop offset="0%" stop-color="#fff" stop-opacity="0.85"/>
                    <stop offset="60%" stop-color="#536DFE" stop-opacity="0.95"/>
                    <stop offset="100%" stop-color="#1A237E" stop-opacity="1"/>
                </radialGradient>
                <radialGradient id="gradGreen" cx="50%" cy="60%" r="80%">
                    <stop offset="0%" stop-color="#fff" stop-opacity="0.85"/>
                    <stop offset="60%" stop-color="#00C853" stop-opacity="0.95"/>
                    <stop offset="100%" stop-color="#1B5E20" stop-opacity="1"/>
                </radialGradient>
                <radialGradient id="gradOrange" cx="50%" cy="60%" r="80%">
                    <stop offset="0%" stop-color="#fff" stop-opacity="0.85"/>
                    <stop offset="60%" stop-color="#FF6D00" stop-opacity="0.95"/>
                    <stop offset="100%" stop-color="#E65100" stop-opacity="1"/>
                </radialGradient>
            `;
            svg.insertBefore(defs, svg.firstChild);
            // Troca as cores das fatias para os gradientes
            var paths = svg.querySelectorAll('g.trace.pie > path');
            var gradIds = ['gradYellow','gradCyan','gradRed','gradBlue','gradGreen','gradOrange'];
            for (var i = 0; i < paths.length; i++) {
                if (paths[i]) paths[i].setAttribute('fill', 'url(#'+gradIds[i%gradIds.length]+')');
            }
        }
    }, 400);
}
