# CODIGO

*Importando xlwings* 

Importando a biblioteca xlwings para interagir com o Excel.

![image](https://github.com/user-attachments/assets/4b5443d8-2ade-48e0-af92-d74c15160d1f)

*Range_values* 

O "@xw.func" registra a função como uma função personalizada do Excel, (intervalo de valores a serem somados), criteria_range (intervalo de valores a serem comparados com o critério), e criteria (critério para a soma).

![image](https://github.com/user-attachments/assets/4c3466e6-97fc-4592-9cf0-7c580186188a)

*Verificação de tamanho*

![image](https://github.com/user-attachments/assets/b360e618-135b-4049-9341-3a1f96616d20)

*Iteração e soma* 

Itera sobre os valores e aplica o critério, somando os valores que atendem ao critério.

![image](https://github.com/user-attachments/assets/6a2014ab-7aa0-483c-8c53-c77360cfc118)

# EXEMPLO

 No excel a tabela seria somada do "A1 ao A10" e "B1 ao B10"
 
 =somase(A1:A10, B1:B10, "Critério")
