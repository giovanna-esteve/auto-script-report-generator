from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import pandas as pd
from datetime import datetime

def mais_lidos_da_semana_no_mundo():
    service = Service(ChromeDriverManager().install())
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Chrome(service=service, options=options)

    website = "https://www.goodreads.com/book/most_read?category=all&country=all&duration=w"
    driver.get(website)

    ##
    livros = driver.find_elements(by="xpath", value='//tr[@itemtype="http://schema.org/Book"]')

    numeros = []
    titulos = []
    autores = []
    leitores = []

    for livro in livros:
        numero = livro.find_element(by="xpath", value='./td').text
        titulo = livro.find_elements(by="xpath", value='./td[3]//span[@itemprop="name"]')[0].text
        autor = livro.find_elements(by="xpath", value='./td[3]//span[@itemprop="name"]')[1].text
        quantidade = livro.find_element(by="xpath", value='./td[3]/span[@class="greyText statistic"]').text

        qnt = quantidade.split(' ')
        numeros.append(int(numero))
        titulos.append(titulo)
        autores.append(autor)
        leitores.append(int(qnt[0].replace(",", "")))

    numeros = numeros[:20]
    titulos = titulos[:20]
    autores = autores[:20]
    leitores = leitores[:20]

    my_dict = {'Numero': numeros, 'Titulo': titulos, 'Autor(a)': autores, 'Leitores': leitores}
    ##

    df = pd.DataFrame(my_dict)
    #now = datetime.now()
    #data = now.strftime("%d%m%Y")
    #df.to_excel(f'{data}_world_week.xlsx')

    driver.quit()
    return df

def mais_lidos_da_semana_no_brasil():
    service = Service(ChromeDriverManager().install())
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Chrome(service=service, options=options)

    website = "https://www.goodreads.com/book/most_read?utf8=%E2%9C%93&country=BR&duration=w"
    driver.get(website)

    ##
    livros = driver.find_elements(by="xpath", value='//tr[@itemtype="http://schema.org/Book"]')

    numeros = []
    titulos = []
    autores = []
    leitores = []

    for livro in livros:
        numero = livro.find_element(by="xpath", value='./td').text
        titulo = livro.find_elements(by="xpath", value='./td[3]//span[@itemprop="name"]')[0].text
        autor = livro.find_elements(by="xpath", value='./td[3]//span[@itemprop="name"]')[1].text
        quantidade = livro.find_element(by="xpath", value='./td[3]/span[@class="greyText statistic"]').text
        qnt = quantidade.split(' ')
        numeros.append(int(numero))
        titulos.append(titulo)
        autores.append(autor)
        leitores.append(int(qnt[0]))

    my_dict = {'Numero':numeros, 'Titulo':titulos, 'Autor(a)':autores, 'Leitores':leitores}
    ##

    df = pd.DataFrame(my_dict)
    #now = datetime.now()
    #data = now.strftime("%d%m%Y")
    #df.to_excel(f'{data}_brazil_week.xlsx')

    driver.quit()
    return df

def pivot_table_brasil_mundo(df_world, df_brazil):
    df_world = df_world[['Titulo', 'Leitores', 'Numero']]
    list_world = df_world.values.tolist()

    df_brazil = df_brazil[['Titulo', 'Leitores', 'Numero']]
    titles_brazil = df_brazil['Titulo'].values.tolist()
    list_brazil = df_brazil.values.tolist()

    local = []
    livros = []
    leitores = []
    numeros = []

    for a, b, c in list_world:
        if a in titles_brazil:
            titulo = a.split('(')[0]

            local.append("Mundo")
            livros.append(titulo)
            leitores.append(b)
            numeros.append(c)

            local.append("Brasil")
            livros.append(titulo)
            index = titles_brazil.index(a)
            leitores.append(list_brazil[index][1])
            numeros.append(list_brazil[index][2])

    my_dict = {'Local': local, 'Livro': livros, 'Leitores': leitores, 'Numeros': numeros}
    df = pd.DataFrame(my_dict)
    pivot_table_completo = df.pivot_table(index='Local', columns='Livro', values='Leitores', aggfunc='sum')

    local = []
    livros = []
    leitores = []
    numeros = []
    for a, b, c in list_world:
        if a in titles_brazil:
            titulo = a.split('(')[0]

            local.append("Mundo")
            livros.append(titulo)
            leitores.append(b)
            numeros.append(c)

    my_dict = {'Local': local, 'Livro': livros, 'Leitores': leitores, 'Numeros': numeros}
    df = pd.DataFrame(my_dict)
    pivot_table_mundo = df.pivot_table(index='Local', columns='Livro', values='Leitores', aggfunc='sum')

    now = datetime.now()
    data = now.strftime("%d%m%Y")
    file_name = f'{data}_relatorio.xlsx'
    with pd.ExcelWriter(file_name) as writer:
        pivot_table_completo.to_excel(writer, sheet_name="graficos", startrow=2)
        pivot_table_mundo.to_excel(writer, sheet_name="pivot_table", startrow=2)

    return file_name
