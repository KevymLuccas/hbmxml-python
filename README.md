# HBM XML - Automação de Download de NFe

**HBM XML** é um sistema em Python com interface gráfica que automatiza o processo de download de Notas Fiscais Eletrônicas (NFe) a partir do site da Receita Federal. Ele foi projetado para facilitar esse processo que, normalmente, exige vários cliques manuais.

## ⚙️ Funcionalidades

* Interface intuitiva com PyQt5
* **Processamento em lote de múltiplas planilhas Excel**
* Importação e exportação de chaves de NFe via Excel
* Gravação de posições dos cliques para automação
* Download automático dos XMLs (**captcha deve ser resolvido manualmente ou por meios externos**)
* **Sistema de retry automático (2 tentativas por NFe)**
* **Tela de bloqueio durante o processo (F11 para emergência)**
* Log detalhado da execução

---

## 🖥️ Requisitos

* Python 3.11+
* PyQt5
* pandas
* pyautogui
* pygetwindow
* openpyxl
* Pillow (para criar logo padrão)

Você pode instalar as dependências com:

```bash
pip install -r requirements.txt
```

---

## 🚀 Como usar

### 1. Primeira Execução (Configuração de Cliques)

Na **primeira vez** que você rodar o programa, ele vai precisar que você **grave as posições dos cliques** na tela do navegador da Receita.

#### Etapas:

1. Execute o programa:

   ```bash
   python hbmxml.py
   ```
2. Clique em **"Baixar XMLs"**.
3. O sistema abrirá o site da Receita Federal.
4. Para **cada passo indicado na interface**, siga os passos abaixo:

   * **ARRASTE a janela do software HBM XML para a coordenada do clique que você deseja gravar.**
   * **CLIQUE sobre a janela do software**, e **não no navegador**.
   * Depois, vá até o navegador e realize a ação normalmente (ex: resolver captcha).
5. Após clicar, a posição será automaticamente registrada no passo correspondente.
6. Repita esse processo até completar os **7 passos**.

As coordenadas gravadas serão usadas na automação e **salvas localmente** tambem serão salvas e usadas automaticamente nas próximas execuções.

---

### 2. Uso Normal (após configuração)

1. Insira manualmente as chaves de NFe no campo superior ou use o botão **"Importar Planilha"**.
2. Clique em **"Baixar XMLs"**.
3. Resolva o captcha quando solicitado.
4. O programa cuidará do resto, clicando automaticamente em todos os passos do processo para cada NFe.

---

## 📄 Explicação dos 7 Passos Gravados

1. Campo da chave da NFe
2. Campo de captcha
3. Botão “Continuar”
4. Botão “Download do Documento”
5. Botão “OK” do popup de confirmação
6. Botão “Nova Consulta”
7. Intervalo entre NFes (espera automática)

---

## 📤 Importação e Exportação

* Você pode importar uma planilha Excel (`.xlsx`) com até 500 chaves de NFe.
* Também pode exportar todas as chaves cadastradas em uma planilha para registro.

---

## 📝 Logs

Durante a execução, um log completo é gerado e exibido na interface e também salvo no arquivo `hbm_xml.log`.

---

## ❗ Observações

* O programa **não resolve captcha automaticamente**. Ele **apenas clica no campo do captcha** como parte do fluxo.
  Se o site exigir uma verificação (ex: imagens, checkbox ou desafios), **o processo não continuará automaticamente**.
* Se você mudar o layout do site da Receita ou **redimensionar/mover os elementos na tela**, será necessário **regravar as posições dos cliques**.
* O sistema depende das **posições exatas** salvas durante a gravação. Mudanças de resolução de tela ou múltiplos monitores podem afetar a automação.

---

## 📦 Código Aberto

Este projeto é open-source. Você pode modificar e distribuir conforme suas necessidades. Pull requests são bem-vindos!

---

## 📧 Suporte

Para dúvidas ou sugestões, abra uma issue neste repositório ou entre em contato comigo.

fantomstore.com.br
