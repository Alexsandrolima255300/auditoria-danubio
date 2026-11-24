# ALEX FREIGHT AUDIT

Sistema de auditoria de fretes da transportadora Danubio.

## Regra inicial
- Frete por peso: peso em kg x R$ 0,70.
- Frete por nota: valor da NF x 1,5%.
- Valor esperado: o maior entre os dois cálculos, respeitando o mínimo de R$ 100,00.

## Regra de elegibilidade
Gerar divergência somente quando o tomador do serviço for BRASIL ENGRENAGENS E CORRENTES IMPORTADORA LTDA, CNPJ 39.860.057/0001-04.

## E-mails
Monitorar mensagens do Siimp Sistemas e pesquisar também a palavra Danubio.

A integração com Outlook, leitura de anexos e persistência ainda precisam ser implementadas.

## Importante: ICMS
O ICMS não entra no cálculo nem na comparação da auditoria. Comparar o valor esperado (maior entre R$ 100,00, peso x R$ 0,70 e NF x 1,5%) exclusivamente com o valor do frete cobrado sem ICMS. Não usar o total da prestação quando ele incluir ICMS.
