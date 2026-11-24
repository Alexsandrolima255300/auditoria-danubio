# ALEX FREIGHT AUDIT

Auditoria de fretes da transportadora Danubio.

## Calculo
- Peso x R$ 0,70.
- Valor da NF x 1,5%.
- Referencia: maior entre R$ 100,00 e os dois calculos.

## Aceitacao
Auditar somente CT-e cujo tomador seja BRASIL ENGRENAGENS E CORRENTES IMPORTADORA LTDA, CNPJ 39.860.057/0001-04.
Para esse tomador, frete cobrado de ate R$ 100,00 (inclusive) e aceito sem divergencia.
Acima de R$ 100,00, comparar com a referencia; diferencas para mais ou para menos geram alerta.
Outros tomadores ficam fora do escopo.

## ICMS
Desconsiderar totalmente o ICMS. Usar somente o valor do frete cobrado, sem ICMS. Nao calcular nem apontar divergencias relacionadas ao imposto.

## Integracoes
Monitoramento Outlook, leitura de anexos e persistencia ainda precisam ser implementados.
