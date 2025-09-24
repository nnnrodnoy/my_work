import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, Circle, Rectangle
import numpy as np

def create_correct_arduino_schema():
    fig, ax = plt.subplots(1, 1, figsize=(12, 8))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 8)
    ax.set_aspect('equal')
    
    # Стиль оформления
    plt.rcParams['font.family'] = 'DejaVu Sans'
    plt.rcParams['font.size'] = 10
    
    # Arduino Board
    arduino = FancyBboxPatch((1, 3), 2, 4, boxstyle="round,pad=0.1", 
                           facecolor='#007ACC', edgecolor='black', linewidth=2)
    ax.add_patch(arduino)
    ax.text(2, 6, 'ARDUINO\nUNO', ha='center', va='center', 
            color='white', weight='bold', fontsize=12)
    
    # Правильные пины согласно схеме
    # +5V pin
    ax.plot([3, 3.2], [6.5, 6.5], 'k-', linewidth=2)
    ax.text(3.5, 6.5, '+5V', ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.3", facecolor='red', color='white'))
    
    # Digital pins
    ax.plot([0.8, 1], [5.5, 5.5], 'k-', linewidth=2)  # D2
    ax.text(0.5, 5.5, 'D2', ha='center', va='center', 
            bbox=dict(boxstyle="round,pad=0.3", facecolor='lightgray'))
    
    ax.plot([0.8, 1], [4.5, 4.5], 'k-', linewidth=2)  # D13
    ax.text(0.5, 4.5, 'D13', ha='center', va='center', 
            bbox=dict(boxstyle="round,pad=0.3", facecolor='lightgray'))
    
    # GND pins
    ax.plot([3, 3.2], [3.5, 3.5], 'k-', linewidth=2)
    ax.text(3.5, 3.5, 'GND', ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.3", facecolor='black', color='white'))
    
    # Components
    # Button (правильное подключение)
    button = Circle((5, 6), 0.4, facecolor='lightblue', edgecolor='black')
    ax.add_patch(button)
    ax.text(5, 6, 'КНОПКА', ha='center', va='center', fontsize=9)
    
    # LED (правильное подключение)
    # Рисуем светодиод правильно - треугольник и линии
    led_triangle = plt.Polygon([[7, 4], [7.5, 4.5], [7, 5]], 
                              facecolor='red', edgecolor='black')
    ax.add_patch(led_triangle)
    # Добавляем линии светодиода
    ax.plot([7, 7], [3.8, 4], 'k-', linewidth=2)  # катод
    ax.plot([7, 7], [5, 5.2], 'k-', linewidth=2)  # анод
    ax.text(7.8, 4.5, 'СВЕТОДИОД', ha='center', va='center', fontsize=9)
    # Обозначаем анод и катод
    ax.text(7.2, 5.1, '+', ha='center', va='center', fontsize=8, color='green')
    ax.text(7.2, 3.9, '-', ha='center', va='center', fontsize=8, color='red')
    
    # Resistors
    # Резистор 10kΩ для кнопки (pull-down)
    resistor1 = Rectangle((4.5, 4.8), 0.2, 1.4, facecolor='brown', edgecolor='black')
    ax.add_patch(resistor1)
    ax.text(4.2, 5.5, '10kΩ', ha='center', va='center', rotation=90, fontsize=8)
    
    # Резистор 220Ω для светодиода
    resistor2 = Rectangle((6.5, 3.8), 1.4, 0.2, facecolor='brown', edgecolor='black')
    ax.add_patch(resistor2)
    ax.text(7.2, 3.5, '220Ω', ha='center', va='center', fontsize=8)
    
    # Правильные соединения
    
    # Кнопка: +5V -> Кнопка -> D2 -> Резистор 10kΩ -> GND
    ax.plot([3.2, 4.6], [6.5, 6.5], 'r-', linewidth=2)  # +5V to button
    ax.plot([5.4, 5.4], [6.5, 5.5], 'r-', linewidth=2)  # button to D2 line
    ax.plot([1, 4.5], [5.5, 5.5], 'r-', linewidth=2)    # D2 to resistor area
    ax.plot([4.5, 4.5], [5.5, 4.8], 'r-', linewidth=2)  # down to resistor
    ax.plot([4.5, 3.2], [4.2, 4.2], 'r-', linewidth=2)  # resistor to GND area
    ax.plot([3.2, 3.2], [4.2, 3.5], 'r-', linewidth=2)  # down to GND
    
    # Светодиод: D13 -> Резистор 220Ω -> Светодиод(анод+) -> Светодиод(катод-) -> GND
    ax.plot([1, 6.5], [4.5, 4.5], 'g-', linewidth=2)    # D13 to resistor
    ax.plot([7.9, 7.9], [4.5, 4.5], 'g-', linewidth=2)  # resistor to LED anode
    ax.plot([7.9, 7.9], [4.5, 4], 'g-', linewidth=2)    # to LED anode connection
    ax.plot([7, 7], [5.2, 5.2], 'g-', linewidth=2)      # from LED cathode
    ax.plot([7, 3.2], [5.2, 5.2], 'g-', linewidth=2)    # LED cathode to GND area
    ax.plot([3.2, 3.2], [5.2, 3.5], 'g-', linewidth=2)  # down to GND
    
    # Labels и пояснения
    ax.text(5, 2.5, 'СХЕМА ПОДКЛЮЧЕНИЯ', ha='center', va='center', 
            fontsize=16, weight='bold', color='darkgreen')
    
    # Легенда
    ax.text(8, 7, 'Обозначения:', ha='left', va='center', weight='bold')
    ax.plot([8, 8.5], [6.7, 6.7], 'r-', linewidth=2)
    ax.text(8.6, 6.7, '- Цепь кнопки', ha='left', va='center', fontsize=9)
    ax.plot([8, 8.5], [6.4, 6.4], 'g-', linewidth=2)
    ax.text(8.6, 6.4, '- Цепь светодиода', ha='left', va='center', fontsize=9)
    
    ax.axis('off')
    plt.tight_layout()
    plt.savefig('correct_arduino_schema.png', dpi=300, bbox_inches='tight', 
                facecolor='lightcyan', edgecolor='none')
    plt.show()

create_correct_arduino_schema()