import xlwings as xw

@xw.func
def somase(range_values, criteria_range, criteria):

    # Verifica se os intervalos são do mesmo tamanho
    if len(range_values) != len(criteria_range):
        raise ValueError("Os intervalos devem ter o mesmo tamanho")
    
    # Começa a soma
    soma = 0
    
    # Junta os valores e aplica o critério
    for value, criterion in zip(range_values, criteria_range):
        if criterion == criteria:
            soma += value
    
    return soma