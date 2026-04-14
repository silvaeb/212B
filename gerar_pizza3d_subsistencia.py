import matplotlib.pyplot as plt
from matplotlib import cm
import numpy as np
import os

# Dados do painel Subsistência
labels = [
    'Crédito disponível (UGR 160504)',
    'Total Descentralizado',
    'Crédito disponível (UGE restantes)',
    'Empenhado',
    'Liquidado',
    'Pago'
]
values = [
    671096423.94,
    570627173.90,
    377201898.59,
    193425275.31,
    1798618.69,
    444172.06
]

# Normalizar valores para porcentagem
values = np.array(values)
percent = values / values.sum() * 100

# Cores customizadas (parecidas com Excel)
colors = [
    '#FFD600',  # amarelo
    '#FF9800',  # laranja
    '#90A4AE',  # cinza
    '#1976D2',  # azul
    '#43A047',  # verde
    '#E53935'   # vermelho
]

fig = plt.figure(figsize=(8, 7), dpi=120)  # altura aumentada
ax = fig.add_subplot(111, projection='3d')

# Parâmetros do 3D
explode = [0.08]*len(values)
startangle = 210

# Função para desenhar fatias 3D
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.patches import Circle
from matplotlib import colors as mcolors

def draw_3d_pie(ax, values, colors, explode, labels, percent, startangle=0):
    total = sum(values)
    theta1 = startangle
    for i, (v, c, e) in enumerate(zip(values, colors, explode)):
        theta2 = theta1 + v/total*360
        x = [0] + np.cos(np.linspace(np.deg2rad(theta1), np.deg2rad(theta2), 30)).tolist()
        y = [0] + np.sin(np.linspace(np.deg2rad(theta1), np.deg2rad(theta2), 30)).tolist()
        z = np.zeros_like(x)
        # Explode
        xm = np.mean(x)
        ym = np.mean(y)
        x = np.array(x) + e*xm
        y = np.array(y) + e*ym
        # Face
        ax.plot_trisurf(x, y, z, color=c, shade=True, alpha=0.98, linewidth=0.2, antialiased=True)
        # Borda
        ax.plot(x, y, z, color='k', lw=0.7, alpha=0.7)
        # Sombra
        ax.plot(x, y, z-0.08, color='k', lw=2, alpha=0.08)
        # Rótulo (percentual + descrição)
        ang = (theta1+theta2)/2
        rx = 1.25*np.cos(np.deg2rad(ang)) + e*np.cos(np.deg2rad(ang))
        ry = 1.25*np.sin(np.deg2rad(ang)) + e*np.sin(np.deg2rad(ang))
        ax.text(rx, ry, 0.13, f'{labels[i]}\n{percent[i]:.1f}%', color=c, fontsize=11, ha='center', va='center', weight='bold', bbox=dict(facecolor='white', alpha=0.7, edgecolor='none', boxstyle='round,pad=0.2'))
        theta1 = theta2

# Limpar fundo
ax.set_facecolor('none')
fig.patch.set_alpha(0)

draw_3d_pie(ax, values, colors, explode, labels, percent, startangle=startangle)

ax.view_init(elev=32, azim=120)  # leve inclinação maior
# Deslocar para cima e direita
ax.set_xlim(-1.1, 1.7)
ax.set_ylim(-1.1, 1.1)
draw_3d_pie(ax, values, colors, explode, labels, percent, startangle=startangle)

# Remover eixos
ax.set_axis_off()

# Salvar imagem com fundo transparente
output_path = os.path.join('static', 'img', 'pizza3d_subsistencia.png')
plt.savefig(output_path, transparent=True, bbox_inches='tight', pad_inches=0.05)
plt.close(fig)
print(f'Gráfico salvo em {output_path}')
