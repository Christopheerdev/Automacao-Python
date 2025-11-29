#🚀 Entrar na planilha do Excel e iniciar automação
import os
import openpyxl
import pyperclip
import pyautogui
from time import sleep

pyautogui.FAILSAFE = False  # Evita parada automática ao mover o mouse para o canto da tela

print("🚀 Iniciando automação de cadastro de produtos...")
sleep(1)

# Abre o Chrome
pyautogui.press('win')
pyautogui.write('chrome', interval=0.1)
pyautogui.press('enter')
sleep(2)

# Abre o site
pyautogui.write('https://cadastro-produtos-devaprender.netlify.app/')
pyautogui.press('enter')
sleep(6)

# Tenta trazer o Chrome para frente
try:
    chrome_windows = pyautogui.getWindowsWithTitle("Chrome")
    if chrome_windows:
        chrome_windows[0].activate()
        print("🌐 Chrome em foco!")
    else:
        print("⚠️ Chrome não encontrado, focando manualmente...")
        pyautogui.hotkey('alt', 'tab')
except Exception as e:
    print("⚠️ Falha ao focar Chrome:", e)
    pyautogui.hotkey('alt', 'tab')

sleep(2)
pyautogui.alert("Clique em OK quando o site estiver totalmente carregado e visível.")

# Caminho absoluto da planilha
base_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(base_dir, "produtos_ficticios.xlsx")

print("📂 Caminho completo da planilha:", file_path)

# Carregar planilha
try:
    workbook = openpyxl.load_workbook(file_path)
    sheet_produtos = workbook["Produtos"]
    print("✅ Planilha carregada com sucesso!")
except Exception as e:
    print(f"❌ Erro ao abrir planilha: {e}")
    exit()

# Loop principal
for i, linha in enumerate(sheet_produtos.iter_rows(min_row=2, values_only=True), start=1):
    nome_produto, descricao_produto, categoria, codigo_ncm, peso, dimensoes, preco, estoque, validade, cor, tamanho, material, fabricante, pais_origem, observacoes, codigo_barras, local_estoque = linha

    print(f"\n📦 Cadastrando produto {i}: {nome_produto}")

    # Campo Nome
    pyperclip.copy(nome_produto)
    pyautogui.moveTo(154,175, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Campo Descrição
    pyperclip.copy(descricao_produto)
    pyautogui.moveTo(147,267, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Categoria
    pyperclip.copy(categoria)
    pyautogui.moveTo(143,393, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Código NCM
    pyperclip.copy(codigo_ncm)
    pyautogui.moveTo(142,479, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')
    # Peso
    pyperclip.copy(peso)
    pyautogui.moveTo(148,568, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Dimensões
    pyperclip.copy(dimensoes)
    pyautogui.moveTo(146,652, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Próximo
    pyautogui.moveTo(147,700, duration=0.5)
    pyautogui.click()
    sleep(4)

    # Preço
    pyperclip.copy(preco)
    pyautogui.moveTo(136,199, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Estoque
    pyperclip.copy(estoque)
    pyautogui.moveTo(129,286, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Validade
    pyperclip.copy(validade)
    pyautogui.moveTo(132,373, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Cor
    pyperclip.copy(cor)
    pyautogui.moveTo(134,458, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Tamanho
    pyautogui.moveTo(198,542, duration=0.5)
    pyautogui.click()
    if tamanho == 'Pequeno':
        pyautogui.click(178,577, duration=0.5)
    elif tamanho == 'Medio':
        pyautogui.click(147,606, duration=0.5)
    else:
        pyautogui.click(144,638, duration=0.5)

    # Material
    pyperclip.copy(material)
    pyautogui.moveTo(136,628, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Próxima página
    pyautogui.moveTo(152,682, duration=0.5)
    pyautogui.click()
    sleep(4)

    # Fabricante
    pyperclip.copy(fabricante)
    pyautogui.moveTo(231,218, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # País de origem
    pyperclip.copy(pais_origem)
    pyautogui.moveTo(132,305, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Observações
    pyperclip.copy(observacoes)
    pyautogui.moveTo(138,390, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Código de barras
    pyperclip.copy(codigo_barras)
    pyautogui.moveTo(131,526, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Local de estoque
    pyperclip.copy(local_estoque)
    pyautogui.moveTo(125,613, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')

    # Concluir
    pyautogui.moveTo(142,669, duration=0.5)
    pyautogui.click()
    sleep(4)

    # OK
    pyautogui.moveTo(850,185, duration=0.5)
    pyautogui.click()
    sleep(3)

    #Botão finalizar
    pyautogui.click(709,438, duration=0.5)
    


    print(f"✅ Produto {i} cadastrado com sucesso!")

    print(f"\n🎯 Processo finalizado com sucesso! Todos os produtos foram cadastrados.")