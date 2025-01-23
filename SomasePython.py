import xlwings as xw

@xw.func
def somase(range_values, criteria_range, criteria):

    # Verifica se os intervalos têm o mesmo tamanho
    if len(range_values) != len(criteria_range):
        raise ValueError("Os intervalos devem ter o mesmo tamanho")
    
    # Inicializa a soma
    soma = 0
    
    # Itera sobre os valores e aplica o critério
    for value, criterion in zip(range_values, criteria_range):
        if criterion == criteria:
            soma += value
    
    return soma

# No Excel ficaria como se somasse como o exemplo abaixo:

# =somase(A1:A10, B1:B10, "Critério")