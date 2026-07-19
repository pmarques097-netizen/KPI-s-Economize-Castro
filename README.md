# Rede Economize — Metas por Comprador (Corrigido)

## Correção aplicada

Foi corrigido o erro:

`NameError: name 'obter_meta_comprador' is not defined`

A causa era a ordem das funções no `app.py`: o painel tentava calcular as
visões antes de carregar o cadastro e as metas individuais dos compradores.

Agora a ordem é:

1. cadastro de compradores;
2. mapa de compradores por classificação;
3. metas por comprador;
4. motor de cálculo das visões;
5. interface do sistema.

## Funcionalidades preservadas

- Cadastro e alteração de compradores
- Comprador por Classificação Principal
- Metas individuais por comprador e período
- Metas gerais
- Banco de dados
- Atualização mensal
- Ruptura automática
- Visões e premiações dinâmicas

## Execução

Feche a janela antiga do sistema e execute:

`EXECUTAR_ECONOMIZE_V2.bat`

A aplicação abre em:

`http://localhost:8502`
