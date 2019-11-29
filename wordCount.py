import xlsxwriter
import xlrd
import sys


class Resultado ():
    def __init__(self, idR, fechaR, tweetR, favR, rtR, palabrasR, institucionR):
        self.idR=idR
        self.fechaR = fechaR
        self.tweetR = tweetR
        self.favR = favR
        self.rtR = rtR
        self.palabrasR = palabrasR
        self.institucionR = institucionR        


def myFunc(e):
  return e['total']
        

def freq(str, workb): 
    palabrasTotal = []
    # break the string into list of words  
    str = str.split()          
    str2 = [] 
  
    # loop till string values present in list str 
    for i in str:              
        
        # checking for the duplicacy 
        if i not in str2: 
  
            # insert value in str2 
            str2.append(i)  
              
    for i in range(0, len(str2)): 
  
        # count the frequency of each word(present  
        # in str2) in str and print 
        # print('Frequency of', str2[i], 'is :', str.count(str2[i]))       
        palabrasTotal.append({"palabra": str2[i], "total": str.count(str2[i])})

    palabrasTotal.sort(key=myFunc, reverse=True)
    
    workbook = xlsxwriter.Workbook(workb+".xlsx")
    worksheet = workbook.add_worksheet()
    row = 0
    for p in palabrasTotal:

        worksheet.write_string(row, 0, p["palabra"])
        worksheet.write_number(row, 1, p["total"])        

        row += 1

    workbook.close()    
    print("Excel file ready")

palabras = ["mujeres", "violencia", "violencia de género", "feminicidio", "personas privadas de su libertad", "derechos humanos", "centros penitenciarios", "adolescentes", "alicia leal", "alejandro encinas", "migración", "migratoria", "migrantes", "equidad", "reincerción social", "víctima", "secretaría de gobernación", "segob", "niñas", "derechos", "violación", "protección","puerta violeta", "género", "mexicanas", "igualdad", "inmujeres", "a_encinas_r", "conavim", "indígena", "comisión nacional de derechos humanos", "cndh","juecez", "juez", "policía", "secretaría", "tribunal", "secretaría", "salud", "política", "ley", "reforma", "sindicato", "deporte", "trabajo", "cultura", "trabajador", "pobresa", "educación", "maestro", "qepd", "deceso", "inauguración", "ciudadania", "fallecimiento", "gratitud", "reconocimiento", "robo", "secuestro", "homicidio", "economía", "ddhh", "conapred", "inami", "instituto nacional de migración", "gobernador", "gobernación", "dmg", "secretaría de la función pública", "conapo", "consejo nacional de población", "snte", "sindicato nacional de trabajadores de la educación", "conavim", "comisión nacional para prevenir y erradicar la violencia contra las mujeres","Infonavit", "instituto del fondo nacional de la vivienda para los trabajadores", "sre", "secretaría de relaciones exteriores", "inafed", "instituto nacional para el federalismo y el desarrollo municipal", "banavim", "banco nacional de datos e información", "scjn", "suprema corte de justicia de la nación", "ammje", "asociación mexicana de mujeres jefas de empresa", "comar", "comisión mexicana de ayuda a refugiados", "was", "world association for sexual health", "imss", "Instituto mexicano del seguro social", "protección civil", "senadomexicano", "dif", "entrevista", "inaugurar", "fiscales", "fiscalía", "jóvenes construyendo el futuro", "seminario", "legislativo", "judicial", "revolución", "estados", "Aaguascalientes", "baja california", "bj", "baja california sur", "bjs", "campeche", "cdmx", "chiapas", "chihuahua", "coahuila" "colima", "durango", "estado de méxico", "guanajuato", "guerrero", "hidalgo", "jalisco", "michoacán", "morelos", "nayarit", "nuevo léon", "oaxaca", "puebla", "querétaro", "quintana roo", "san luis potosi", "san luis", "sinaloa", "sonora", "tabasco", "tamaulipas", "tlaxcala", "veracruz", "yucatán", "zacatecas", "participacion", "participación", "toma de protesta", "legislador", "legislación", "cumpleaños", "presidente", "presidenc", "evoespueblo", "felicitación", "felicidades", "felicitaciones", "Ricar_peralta", "embajada", "embajador", "estrategia", "impunidad", "condolencias", "estudio", "estudiante", "luto", "violentada", "menospreciad", "marginad", "homenaje", "huachicoleo", "delito", "explosión", "felicita", "felicito", "agradezco", "agradecimiento", "agradecer", "agradece", "naciones unidas", "onu", "refugiad", "refugio", "senado", "senado de la republica", "cooperacion","cooperación", "centros de justicia", "secretaría de la función pública" "sepulcro", "homenaje", "inclusión", "discriminación", "4t", "informe de gobierno", "criminales", "estados unidos", "jóvenes", "niños", "hombres", "reunión", "leer", "lectura", "libro", "derechos", "humanos", "ayotzinapa", "designación", "libertad de expresión", "prensa"]

instituciones = ["conapred", "comisión nacional de derechos humanos", "cndh", "inmujeres", "segob", "secretaria de gobernación", "inami", "instituto nacional de migración", "gobernador", "gobernación", "dmg", "secretaría de la función pública", "conapo", "consejo nacional de población", "snte", "sindicato nacional de trabajadores de la educación", "conavim", "comisión nacional para prevenir y erradicar la violencia contra las mujeres","infonavit", "instituto del fondo nacional de la vivienda para los trabajadores", "sre", "secretaría de relaciones exteriores", "inafed", "instituto nacional para el federalismo y el desarrollo municipal", "banavim", "banco nacional de datos e información", "scjn", "suprema corte de justicia de la nación", "ammje", "asociación mexicana de mujeres jefas de empresa", "comar", "comisión mexicana de ayuda a refugiados", "was", "world association for sexual health", "imss", "Instituto mexicano del seguro social", "protección civil", "senadomexicano", "dif", "naciones unidas", "onu", "senado de la republica", "centros de justicia", "secretaría de la función pública", "comisión de pesca", "cámara de diputado", "cámaras de diputados", "marina","armada", "colegio nacional del notariado mexicano", "gobiernomx", "unam", "presidente municipal", "academiaidh", "embajada", "embajador", "secretaría de seguridad ciudadana", "comisión presidencial para la conmemoración de hechos, procesos y personajes históricos de méxico", "comisión de conmemoraciones", "comisión para la verdad y acceso a la justicia en el caso ayotzinapa", "fgr", "fiscalía general de la república"]


tweets = []
indexs = []
fechas = []
ids = []
contadorRT = []
contadorFav = []
index = 1

stringsote = ""
stringsoteT = ""
stringsoteRT = ""

totalT = []
resultados = []
rts = []


workbook = xlrd.open_workbook("M_OlgaSCordero2.xlsx")
sheet = workbook.sheet_by_index(0)

for i in range(sheet.nrows):    
    ids.append(sheet.cell_value(i, 0)) 
    fechas.append(sheet.cell_value(i, 1))
    tweets.append(sheet.cell_value(i, 2))
    contadorFav.append(sheet.cell_value(i, 4))
    contadorRT.append(sheet.cell_value(i, 5))
    indexs.append(index)
    index = index + 1

j = 0
for tweet in tweets: 

    temp = tweet.replace(",","")
    temp = temp.replace(".","")
    temp = temp.replace(":","")
    temp = temp.replace("!","")
    temp = temp.replace("¡","")
    temp = temp.replace("¿","")
    temp = temp.replace("?","")
    temp = temp.replace("()","")
    temp = temp.replace(")","")
    temp = temp.replace('"',"")
    temp = temp.replace("'","")    

    resultado = ""
    inst = ""
    for palabra in palabras:

        if(palabra in tweet.lower()):
            resultado = resultado + palabra + ","
    
    for institucion in instituciones:

        if(institucion in tweet.lower()):
            inst = inst + institucion + ","
    
    if(tweet.find("RT",0,2)>=0):
        rts.append(Resultado(ids[j],fechas[j], tweet, contadorFav[j], contadorRT[j], resultado, inst))
        totalT.append(Resultado(ids[j],fechas[j], tweet, contadorFav[j], contadorRT[j], resultado, inst))
        stringsoteRT = stringsoteRT + " "+ temp

    else:
        resultados.append(Resultado(ids[j],fechas[j], tweet, contadorFav[j], contadorRT[j], resultado, inst))
        totalT.append(Resultado(ids[j],fechas[j], tweet, contadorFav[j], contadorRT[j], resultado, inst))
        stringsoteT = stringsoteT + " "+ temp
    
        

    stringsote = stringsote + " "+ temp
    j = j + 1

print(len(stringsote))
print(len(stringsoteT))
print(len(stringsoteRT))


#freq(stringsote, "TopPalabrasTotal")
#freq(stringsoteT, "TopPalabrasTweet")
#freq(stringsoteRT, "TopPalabrasRT")


workbook = xlsxwriter.Workbook("TweetCount.xlsx")
worksheet = workbook.add_worksheet()
row = 0
for resultado in resultados:

    worksheet.write_string(row, 0, resultado.idR)
    worksheet.write_string(row, 1, resultado.fechaR)
    worksheet.write_string(row, 2, resultado.tweetR)
    worksheet.write_string(row, 3, resultado.favR)
    worksheet.write_string(row, 4, resultado.rtR)
    worksheet.write_string(row, 5, resultado.palabrasR)
    worksheet.write_string(row, 6, resultado.institucionR)

    row += 1

workbook.close()
print("Excel file ready")


workbook = xlsxwriter.Workbook("TotalCount.xlsx")
worksheet = workbook.add_worksheet()
row = 0
for t in totalT:

    worksheet.write_string(row, 0, t.idR)
    worksheet.write_string(row, 1, t.fechaR)
    worksheet.write_string(row, 2, t.tweetR)
    worksheet.write_string(row, 3, t.favR)
    worksheet.write_string(row, 4, t.rtR)
    worksheet.write_string(row, 5, t.palabrasR)
    worksheet.write_string(row, 6, t.institucionR)

    row += 1

workbook.close()
print("Excel file ready")

workbook = xlsxwriter.Workbook("RTCount.xlsx")
worksheet = workbook.add_worksheet()

row = 0
for rt in rts:

    worksheet.write_string(row, 0, rt.idR)
    worksheet.write_string(row, 1, rt.fechaR)
    worksheet.write_string(row, 2, rt.tweetR)
    worksheet.write_string(row, 3, rt.favR)
    worksheet.write_string(row, 4, rt.rtR)
    worksheet.write_string(row, 5, rt.palabrasR)
    worksheet.write_string(row, 6, rt.institucionR)

    row += 1

workbook.close()
print("Excel file ready")