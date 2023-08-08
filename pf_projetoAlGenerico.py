import random
from openpyxl import Workbook
import time

referencia=[1]*10
tamanhoPopulacaoPeriodo = 10
tamanhoPopulaçãoCurso = 15
#Classe responsavel por gerar o objeto disciplina
#Apresnta as variaveis nomeDisciplina, cargaHoraria, professor, periodo#
class disciplina:
    def __init__(self, nomeDis, cargaHoraria, professor, periodo, restricoes):
        self.nomeDisciplina = nomeDis
        self.cargaHoraria = cargaHoraria
        self.professor = professor
        self.periodo = periodo
        self.restricoes = restricoes

    def encontros(self):
        return int(self.cargaHoraria/30)
    
#Inicia a população com valores 0's para dias sem a disciplian e 1's para dias com a disciplina
#É determinado o numero de dias que tera a disciplina de acordo com a quantidade de encontros (previamente determinada)
# e o tamanho da população previamente passada pelo usuario#
def initPopulacao(dis, tamanhoPopulacao):
    temp = ['']*10
    i = 0
    rr = []
    encontros = dis.encontros()
    quantEncontros = 0
    diasDisponiveis = []
    for k in range(10):
        if((k in dis.restricoes) == False):
            diasDisponiveis.append(k)
    
    while(i!=tamanhoPopulacao):
        teste = []
        while(quantEncontros != encontros):        
            rd = random.choice(diasDisponiveis)     
            if((rd in teste) == False):  
                quantEncontros = quantEncontros + 1
                temp[rd] = 1
                teste.append(rd)
            
        
        for x in range(10):            
            if(temp[x] != 1):                                        
                temp[x] = 0
            
        if(encontros == quantEncontros):            
            rr.append(temp)
            temp = ['']*10
            quantEncontros = 0
            i = i + 1
    return rr

#Determina a aptidão do horario das disciplia
#maior aptidão => Menos choque de horario entre as disciplinas#
def ap2(grade):    
    temp = [0]*10
    for i in range(len(grade)):
        for n in range (len(grade[i])):
            if(grade[i][n] == 1):
                temp[n] = 1
    cont = 0
    for j in range(len(referencia)):
        if(temp[j] == referencia[j]):
            cont +=1
    return cont

def aptidao(grade):
    grades, materias = separa(grade)
    temp = [0]*10
    for i in range(len(grades)):
        for n in range (len(referencia)):
            if(grades[i][n] == 1):
                temp[n] = 1
    cont = 0
    for j in range(len(referencia)):
        if(temp[j] == referencia[j]):
            cont +=1
    return cont

#Divide a matriz que contem a grade e as disciplinas
#e retorna a matriz de grade de horario e a matriz de professor correspondente#
def separa(grade):
    tamanho = int(len(grade)/2) 
    grades = []
    turmas = []
    for i in range(tamanho):
        grades.append(grade[i])
    for n in range(tamanho, len(grade), 1):
        turmas.append(grade[n])
    return grades, turmas

#Realiza a junção entre uma matriz de grade e a matriz de disciplinas#
def junta(grades, turmas):
    juncao = grades
    for i in range(len(turmas)):
        juncao.append(turmas[i])
    return juncao

#Realiza o primerio Crossover entre a grade de duas disciplinas
def crossoverPopulacao(d1, d2):
    pai1 = random.randint(0, len(d1)-2)
    pai2 = random.randint(0, len(d2)-2)
    filho = [d1[pai1], d2[pai2]]
    filho.append(d1[len(d1)-1])
    filho.append(d2[len(d2)-1])
    return filho

#Realiza o crossover entre duas ou mais disciplinas formando assim o horario completo
def cros2(grade, populacao):
    teste = random.randint(0,len(populacao)-2)
    grades, materias = separa(grade)
    grades.append(populacao[teste])
    materias.append(populacao[len(populacao)-1])
    juncao = junta(grades, materias)
    return juncao       

#Metodo responsavel por gerar os periodos
#recebendo um vetor contendo todas as disciplinas de um determinado periodo
#e retornando a matriz com as disciplinas e seus respectivos horarios# 
def periodos(vetorDisciplinas):
    tamanhoAp = 0
    for i in range(len(vetorDisciplinas)):
        tamanhoAp += vetorDisciplinas[i].encontros()
    temp = []
    gradePeriodo = []
    apTemp = 0
    if(len(vetorDisciplinas) == 2):
        while(aptidao(temp) != tamanhoAp):
            temp = gradePeriodo         
            pop1 = initPopulacao(vetorDisciplinas[0], tamanhoPopulacaoPeriodo)    
            pop2 = initPopulacao(vetorDisciplinas[1], tamanhoPopulacaoPeriodo)
            pop1.append(vetorDisciplinas[0])
            pop2.append(vetorDisciplinas[1])
            temp = crossoverPopulacao(pop1,pop2)
        gradePeriodo = temp
        return gradePeriodo
    elif(len(vetorDisciplinas) > 2):
        while(aptidao(temp) != tamanhoAp):
            while(aptidao(temp) != vetorDisciplinas[0].encontros()+vetorDisciplinas[1].encontros()):
                temp = gradePeriodo         
                pop1 = initPopulacao(vetorDisciplinas[0], tamanhoPopulacaoPeriodo)    
                pop2 = initPopulacao(vetorDisciplinas[1], tamanhoPopulacaoPeriodo)
                pop1.append(vetorDisciplinas[0])
                pop2.append(vetorDisciplinas[1])
                temp = crossoverPopulacao(pop1,pop2)
            gradePeriodo = temp
            apTemp = vetorDisciplinas[0].encontros()+vetorDisciplinas[1].encontros()
            for j in range(2, len(vetorDisciplinas), 1):
                gradePeriodo = temp
                while(aptidao(temp) != apTemp + vetorDisciplinas[j].encontros()):
                    temp = gradePeriodo
                    pop = initPopulacao(vetorDisciplinas[j], tamanhoPopulacaoPeriodo)
                    pop.append(vetorDisciplinas[j])
                    temp = cros2(gradePeriodo, pop)
                apTemp += vetorDisciplinas[j].encontros()        
                       
                gradePeriodo = temp
        return gradePeriodo



#Seleciona os professores que apresentam duas ou mais disciplinas no curso e retorna os mesmos#
def profMultiDis(curso):
    dis = []
    for i in range(len(curso)):
        for j in range(int(len(curso[i])/2), len(curso[i]),1):
            for n in range(len(curso)):
                for k in range (int(len(curso[n])/2), len(curso[n]),1):
                    for r in range(len(curso[i][j].professor)):
                        for s in range(len(curso[n][k].professor)):
                            if(curso[i][j].professor[r] == curso[n][k].professor[s] and i != n and curso[n][k].periodo != curso[i][j].periodo):
                                if((curso[i][j].professor in dis) == False):                                    
                                    dis.append(curso[i][j].professor)
    
    return dis

#separa a matriz do curso em duas outras matrizes, uma contendo apenas os horarios e outra contendo apenas as disciplinas respectivas aos horarios 
def separaCurso(curso):
    horarios = []
    dis = []
    for i in range(len(curso)):
        temp1, temp2 = separa(curso[i])
        horarios.append(temp1)
        dis.append(temp2)        
    return horarios, dis

#Determina a aptidão do curso calculando qual o total de encontros que os professores devem ter sem apresentar choque
# e comparado com a real sitação da grade, quando mais proximo de 100, maior será sua aptidão#
def aptidaoCurso(curso):    
    disMult = profMultiDis(curso)                              
    horario, disciplinas = separaCurso(curso)
    aptidaoC = 0
    temp2 = 0

    for i in range(len(disMult)):
        for s in range(len(disMult[i])):
            temp = []
            for n in range(len(disciplinas)):
                for k in range(len(disciplinas[n])):
                    for j in range(len(disciplinas[n][k].professor)):
                        if(disMult[i][s] == disciplinas[n][k].professor[j]):
                            temp.append(horario[n][k])
                            temp2 += disciplinas[n][k].encontros()
            aptidaoC += ap2(temp)
    
    if(temp2 !=0):
        aptidaoFinal = (aptidaoC * 100)/temp2
    else:
        aptidaoFinal = (aptidaoC * 100)/1
    return aptidaoFinal

#Seleciona os individuos com maiores aptidões da população de grades do periodo
# para serem os genitores para as proximas gerações de filhos#
def pais(populacao):
    apt = []

    for i in range(len(populacao)):
        apt.append(aptidaoCurso(populacao[i]))

    aptTotal = sum(apt)
    probabilidade = []

    for n in range(len(apt)):
        probabilidade.append(apt[n]/aptTotal)
    
    pais = []
    for j in range(len(probabilidade)):
        if(probabilidade[n] > 0):
            pais.append(populacao[j])
    return pais

#a partir dos pais selecionados, é escolhido dois randomicamente e os funde, gerando assim um filho, com partes de um pai
# e partes do outro pai#
def descendente(populacao):
    pai1 = populacao[random.randrange(0, len(populacao))]
    pai2 = populacao[random.randrange(0, len(populacao))]
    cross = random.randint(0, 8) #seleciona o ponto de junção
    filho = []

    for i in range(cross):
        filho.append(pai1[i])
    for i in range(cross, 9, 1):
        filho.append(pai2[i])    
    return filho

#Dentre a população seleciona os com as melhores aptidões para sobreviver e os demais são descartados#
def sobreviventes(populacao):
    aux = []
    aux2 = []
    for i in range(len(populacao)-1):
        if(aptidaoCurso(populacao[i]) > aptidaoCurso(aux) and (populacao[i] in aux2) == False):
            aux = populacao[i]
            aux2.append(populacao[i])
    return aux2

#Realiza a mutação em uma filho, para assim não acabar com a evolução do curso
# onde seleciona um periodo aleatoriamente e o refaz#
def xman(cursoFilho, vetorPeriodos):
    cursoTemp = cursoFilho
    periodo = random.randint(0, len(vetorPeriodos)-1)
    periodoX = vetorPeriodos[periodo]
    cursoTemp[periodo] = periodos(periodoX)
    if(aptidaoCurso(cursoFilho) >= aptidaoCurso(cursoTemp)):
        return cursoFilho
    else:
        return cursoTemp

#1º periodo
ip = disciplina('IP', 90, ['Luis','Rene'],1, [0,5]) 
log = disciplina('logica Matematica', 60, ['Gersonilo'],1,[0,5])
ic = disciplina('IC', 30, ['Ryan'],1,[3,8,4,9])
ga = disciplina('GA', 60, ['Normando'],1, [])
c1 = disciplina('C1', 60, ['Marcius'],1, [])

v1 = [ic,ip,log,ga,c1]

#2º Periodo
aL = disciplina('Algebra Linear', 60, ['Gersonilo'], 2,[0,5])
poo = disciplina('POO', 60, ['Igor'],2, [2,7])
aed1 = disciplina('AED 1', 60, ['Igor','Rene'],2, [2,7])
c2 = disciplina('Calculo 2', 60, ['Sansuke'], 2, [])
fisica = disciplina('Fisica 1', 60, ['Wellington'],2, [])

v2 = [poo, aed1,aL,c2, fisica]

#3º Periodo
mD = disciplina('Matematica Discreta', 60, ['Gersonilo'], 3, [0,5])
ingles = disciplina('Ingles', 30, ['Diana'], 3, [])
sisDigitais = disciplina('Sistemas Digitais', 60, ['Helder'], 3, [])
aed2 = disciplina('AED 2', 60, ['Igor'], 3, [2,7])
metCientifica = disciplina('Metodologia Cientifica', 30, ['Leila'], 3, [])
probEst = disciplina('Probabilidade e estatistica', 60, ['Romero'], 3, [])

v3 = [ mD,ingles, sisDigitais, aed2, metCientifica, probEst]

#4º Periodo
paa =disciplina('PAA', 60, ['Alvaro'], 4,[])
arqComp = disciplina('Arquitetura de computadores', 60, ['Helder'], 4,[])
bd = disciplina('BD', 60, ['Priscila'], 4, [])
engSoftware = disciplina('Eng. de Software', 60, ['R. Andrade'], 4, [])
plp = disciplina('PLP',60,['Ryan'], 4, [3,8,4,9])

v4 = [plp,paa,arqComp, bd, engSoftware]

#5º Periodo
sdt = disciplina('Sistemas de informacao e sua Tec.', 60,['Assuero'], 5, [])
redes = disciplina('Redes de computadores', 60, ['Kadna'], 5, [])
teoria = disciplina('Teoria da Computacao', 60, ['Maria'], 5, [])
so = disciplina('SO', 60,['Sergio'], 5, [])
ia = disciplina('IA', 60, ['Tiago'], 5, [])

v5 = [sdt, redes, teoria, so, ia]

#6º Periodo
emp = disciplina('Empreendedorismo', 60, ['Assuero'], 6, [])
cg =  disciplina('Computacao Grafica', 60, ['Icaro'], 6, [])
sisDist = disciplina('Sistemas Distribuidos', 60,['Jean'], 6,[])
rp = disciplina('Reconhecimento de padroes', 60,['Luis'], 6, [0,5])
comp = disciplina('Compiladores', 60, ['Maria'], 6,  [])

v6 = [rp, emp, cg, sisDist,comp]

#7º Periodo
mdsc = disciplina('Modelagem de Dependabilidade de Sistemas Computacionais', 60, ['Jean'], 7, [])
ihc = disciplina('Interação Humano computador', 60, ['R. Andrade'], 7, [])
projeto = disciplina('Projeto de desenvolvimento', 90,['R. Gusmão'], 7, [])
compSo = disciplina('Computadores e Sociedade', 30,['Ryan'], 7,[3,8,4,9])
deps = disciplina('Desenvolvimento e Exec. de projeto de software', 60,['Ryan'], 7, [3,8,4,9])

v7 = [compSo, deps, mdsc, ihc, projeto]

#8º Periodo
ts = disciplina('Teste de software', 60, ['Alvaro'], 8,[])
mpc = disciplina('Metodos de pesquisa em computação', 60,['Kadna'], 8, [])
rn = disciplina('Redes Neurais', 60, ['Luis'], 8,[0,5])
bda = disciplina('BD avançado', 60, ['Priscila'], 8,  [])
am = disciplina('Aprendisagem de maquinas', 60, ['Tiago'], 8, [])

v8 = [rn, ts, mpc,  bda, am]

#9º Periodo
eso = disciplina('Estagio Supervisionado Obrigatorio', 30, ['R. Gusmão'], 9, [])
segRedesPc = disciplina('Segurança de redes de computadores', 60, ['Sergio'], 9, [])
v9 = [eso, segRedesPc]



vetorPeriodos= [v1,v2,v3,v4,v5,v6,v7,v8,v9]
contador = 0

melhorCurso = []
pop = []
populacao = 0
inicio = time.time() # inicio da exec. do codigo

while(populacao != tamanhoPopulaçãoCurso):#inicializa a criação da populaçao de grade (é a parte que mais demora na exec.)
    curso = []
    p1 = periodos(v1)        
    p2 = periodos(v2)        
    p3 = periodos(v3)        
    p4 = periodos(v4)        
    p5 = periodos(v5)        
    p6 = periodos(v6)        
    p7 = periodos(v7)        
    p8 = periodos(v8)        
    p9 = periodos(v9)    
    
    curso = [p1,p2,p3,p4,p5,p6,p7,p8,p9]
    
    pop.append(curso)    
    populacao+=1

for i in range(len(pop)): #dentre a população inicial, determina o individuo(grade de curso) com maior aptidão
    if(aptidaoCurso(pop[i]) > aptidaoCurso(melhorCurso)):
        melhorCurso = pop[i]
geracoes = 0
while(aptidaoCurso(melhorCurso) != 100): #para a exec. apenas quando apresentar um curso com aptidão == 100
   
    for i in range(int(len(pop)/2)):  #Gera filhos e os add a população
        ancedentes = pais(pop)
        filho1 = xman(descendente(ancedentes), vetorPeriodos)
        filho2 = xman(descendente(ancedentes), vetorPeriodos)          
        pop.append(filho1)        
        pop.append(filho2)

    pop = sobreviventes(pop)

    for i in range(len(pop)): #dentre a população com os filhos, determina o individuo(grade de curso) com maior aptidão
        if(melhorCurso != pop[i]):
            if(aptidaoCurso(pop[i]) > aptidaoCurso(melhorCurso)):
                melhorCurso = pop[i]
            else:
                pop.append(melhorCurso)
    geracoes+=1
    print(f"{int(aptidaoCurso(melhorCurso))}% de aptidão")

fim = time.time() #Fim da exec. do codigo

horario, materias = separaCurso(melhorCurso)
print(geracoes,"<-- Numero de gerações")
print(fim - inicio,"<-- Tempo de Execução")



for k in range(len(horario)): #imprime a grade do curso no terminal
    for l in range(len(horario[k])):
        print(horario[k][l], materias[k][l].nomeDisciplina, materias[k][l].professor, materias[k][l].periodo)
        
workbook = Workbook()
sheet = workbook.active
cont = 0
for i in range(1,37,4):#imprime a grade do curso em formato .xlms
    
    sheet[f"A{i+1}"] = "18:30 - 20:10"
    sheet[f"A{i+2}"] = "20:10 - 21:50"

    sheet[f"B{i}"] = "Segunda-Feira"
    sheet[f"C{i}"] = "Terça-Feira"
    sheet[f"D{i}"] = "Quarta-Feira"
    sheet[f"E{i}"] = "Quinta-Feira"
    sheet[f"F{i}"] = "Sexta-Feira"
    sheet[f"A{i}"] = f"{cont+1}º Periodo"

    for n in range(len(horario[cont])):
        for j in range(len(horario[cont][n])):
            if(horario[cont][n][j] == 1 and j == 0):
                sheet[f"B{i+1}"] = f"{materias[cont][n].nomeDisciplina}, {materias[cont][n].professor}"
            elif(horario[cont][n][j] == 1 and j == 1):
                sheet[f"C{i+1}"] = f"{materias[cont][n].nomeDisciplina}, {materias[cont][n].professor}"
            elif(horario[cont][n][j] == 1 and j == 2):
                sheet[f"D{i+1}"] = f"{materias[cont][n].nomeDisciplina}, {materias[cont][n].professor}"
            elif(horario[cont][n][j] == 1 and j == 3):
                sheet[f"E{i+1}"] = f"{materias[cont][n].nomeDisciplina}, {materias[cont][n].professor}"    
            elif(horario[cont][n][j] == 1 and j == 4):
                sheet[f"F{i+1}"] = f"{materias[cont][n].nomeDisciplina}, {materias[cont][n].professor}"
            elif(horario[cont][n][j] == 1 and j == 5):
                sheet[f"B{i+2}"] = f"{materias[cont][n].nomeDisciplina}, {materias[cont][n].professor}"
            elif(horario[cont][n][j] == 1 and j == 6):
                sheet[f"C{i+2}"] = f"{materias[cont][n].nomeDisciplina}, {materias[cont][n].professor}"
            elif(horario[cont][n][j] == 1 and j == 7):
                sheet[f"D{i+2}"] = f"{materias[cont][n].nomeDisciplina}, {materias[cont][n].professor}"
            elif(horario[cont][n][j] == 1 and j == 8):
                sheet[f"E{i+2}"] = f"{materias[cont][n].nomeDisciplina}, {materias[cont][n].professor}"
            elif(horario[cont][n][j] == 1 and j == 9):
                sheet[f"F{i+2}"] = f"{materias[cont][n].nomeDisciplina}, {materias[cont][n].professor}"
    cont+=1  

sheet["A37"] = f"{geracoes} gerações"
sheet["A38"] = f"{int(fim - inicio)} tempo em seg"
workbook.save(filename="grade_de_horario_bcc.xlsx")