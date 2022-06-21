
ValorDecimal = 128000021.82
Tamanho = len(str(ValorDecimal))

ValorInteiro = (str(ValorDecimal)[:(Tamanho-3)])
ValorDecimal = (str(ValorDecimal)[(Tamanho-3)+1:])


print(ValorInteiro)
print(ValorDecimal)
