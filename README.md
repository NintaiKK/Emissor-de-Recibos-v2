# Sistema de Recibos Automatizado
Um aplicativo desktop desenvolvido em Python para geração automatizada de recibos de pagamento e adiantamento, com gestão de funcionários e descontos.

# Funcionalidades Principais
 - Geração de Recibos:

 - Criação de recibos de quitação para pagamentos ou adiantamentos

 - Cálculo automático de valores totais com descontos

 - Geração do valor por extenso em português

 - Exportação para Excel com formatação profissional

### Gestão de Funcionários:

 - Cadastro completo com nome, CPF/CNPJ, salário e valores padrão

 - Edição e exclusão de registros

 - Armazenamento em arquivo XML

# Recursos Avançados:

 - Sistema flexível de descontos (adicionar/remover dinamicamente)

 - Cálculo automático de valores líquidos

 - Pré-visualização em tempo real

 - Suporte a parcelas extras

# Tecnologias Utilizadas
 - Python 3

 - Tkinter (GUI)

 - OpenPyXL (manipulação de Excel)

 - XML (armazenamento de dados)

 - Locale (formatação brasileira)

# Como Usar
Cadastre os funcionários na aba "Gerenciar Funcionários"

Selecione o funcionário e preencha os dados na aba "Emitir Recibo"

Adicione descontos se necessário

Visualize o recibo na aba de pré-visualização

Gere o arquivo Excel com o botão "Gerar Recibo"

# Requisitos
Python 3.x

Bibliotecas: tkinter, openpyxl, xml.etree.ElementTree

# Instalação
```pip install openpyxl```

Clone o repositório e execute:

```python sistema_recibos.py```
# Licença

GNU 3
