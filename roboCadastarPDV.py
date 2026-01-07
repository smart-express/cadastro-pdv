import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import re
from datetime import datetime

nomeArquivo = 'CADASTRARPDV.xlsx'
df = pd.read_excel(nomeArquivo, dtype=str)

async def inputForm(pagina, campo, dado):
    await pagina.locator(f'//label[normalize-space()="{campo}"]/preceding-sibling::input[1]').fill(dado)
    #await pagina.fill(f"input[formcontrolname={campo}]", dado)

async def inputFormClick(pagina, campo):
    await pagina.locator(f'//label[normalize-space()="{campo}"]/preceding-sibling::input[1]').click()
    #await pagina.fill(f"input[formcontrolname={campo}]", dado)
    


def clearNumber(numero: str) -> str:
    numero = numero.strip()
    return numero.replace("+55", "", 1).strip()

def clearCpf(doc: str) -> str:
    return "".join(c for c in doc if c.isdigit())

def limpar_cpf(cpf: str) -> str:
    return cpf.replace(".", "").replace("-", "").zfill(11)

def handleData(data_str: str) -> str:
    for formato in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            data = datetime.strptime(data_str, formato)
            return data.strftime("%d/%m/%Y")
        except ValueError:
            pass

    raise ValueError(f"Formato de data inválido: {data_str}")

async def main():

    async with async_playwright() as p:
        navegador = await p.chromium.launch(headless=False)
        context = await navegador.new_context()
        pagina = await context.new_page()

        await pagina.goto("https://gmpromo.pdvnow.com/#/entrar")

        await pagina.get_by_role("textbox", name="Login").wait_for()
        await pagina.get_by_role("textbox", name="Login").fill('43812401894')

        await pagina.get_by_role("textbox", name="Senha").wait_for()
        await pagina.get_by_role("textbox", name="Senha").fill('12345')

        await pagina.get_by_role("button", name="Fazer Login").click()
        await pagina.get_by_role("link", name="Usuários").click()

        for index, row in df.iterrows():

            await pagina.get_by_role("textbox", name="Pesquisar").fill(limpar_cpf(row["CPF"]))
            await pagina.keyboard.press("Enter")
            await asyncio.sleep(2)

            if not await pagina.get_by_role("cell", name=limpar_cpf(row["CPF"])).is_visible():
                
                await pagina.get_by_role("button", name="Adicionar usuário").click()
                await inputForm(pagina, 'Nome', row['Nome Completo:'])
                await inputForm(pagina, 'E-mail', row['Email'])
                await inputForm(pagina, 'Telefone', clearNumber(row['Número de contato:']))
                await inputForm(pagina, 'Login', row["CPF"])
                await inputForm(pagina, 'Senha', "express1234")
                await inputForm(pagina, 'Projeto', "COMPARTILHADO")
                
                await pagina.get_by_role("listbox").filter(has_text="Superior").get_by_role("combobox").click()
                await pagina.get_by_role("listbox").filter(has_text="Superior").get_by_role("combobox").fill(row["Nome do Lider Responsavel:"])
                await pagina.keyboard.press("Enter")

                await pagina.get_by_role("listbox").filter(has_text="Perfil de acesso").get_by_role("combobox").click()
                await pagina.get_by_role("listbox").filter(has_text="Perfil de acesso").get_by_role("combobox").fill("Express")
                await pagina.keyboard.press("Enter")

                if row['Perfil Smart'] == "EXPRESS":
                    await pagina.get_by_role("listbox").filter(has_text="Empresa").get_by_role("combobox").nth(0).click()
                    await pagina.get_by_role("listbox").filter(has_text="Empresa").get_by_role("combobox").nth(0).fill("Desenvolver")
                    await pagina.keyboard.press("Enter")
                
                if row['Perfil Smart'] == "FREELANCER":
                    await pagina.get_by_role("listbox").filter(has_text="Empresa").get_by_role("combobox").nth(0).click()
                    await pagina.get_by_role("listbox").filter(has_text="Empresa").get_by_role("combobox").nth(0).fill("freelancer")
                    await pagina.keyboard.press("Enter")
                
                if row['Perfil GMP'] == "Flex":
                    await pagina.get_by_role("listbox").filter(has_text="Empresa").get_by_role("combobox").nth(0).click()
                    await pagina.get_by_role("listbox").filter(has_text="Empresa").get_by_role("combobox").nth(0).fill("flex")
                    await pagina.keyboard.press("Enter")
                    
                await pagina.get_by_role("listbox").filter(has_text="Roteiro").get_by_role("combobox").click()
                await pagina.get_by_role("option", name="Sim").click()
                await pagina.keyboard.press("Enter")
                
                await pagina.get_by_role("listbox").filter(has_text="Status").get_by_role("combobox").click()
                await pagina.get_by_role("option", name="Inativo").click()
                await pagina.keyboard.press("Enter")
                
                await pagina.get_by_role("listbox").filter(has_text="Empresa login").get_by_role("combobox").click()
                await pagina.get_by_role("option", name="GM PROMO").click()
                await pagina.keyboard.press("Enter")
                
                await pagina.get_by_role("listbox").filter(has_text="Cargo de campo").get_by_role("combobox").click()
                await pagina.get_by_role("listbox").filter(has_text="Cargo de campo").get_by_role("combobox").fill("Roteiro")
                await pagina.keyboard.press("Enter")
                
                await pagina.get_by_role("link", name="Dados pessoais").click()
                
                await inputForm(pagina, 'Data de nascimento', handleData(row["Data de Nascimento"]))
                await inputForm(pagina, 'CPF', row["CPF"])
                
                if pd.notna(row["CNPJ"]):
                    await inputForm(pagina, 'CNPJ', row["CNPJ"])
                    
                await inputForm(pagina, 'RG', row["Número do documento de identificação"])
                await pagina.get_by_role("listbox").filter(has_text="Sexo").get_by_role("combobox").click()
                await pagina.get_by_role("listbox").filter(has_text="Sexo").get_by_role("combobox").fill(row["Sexo"])
                await pagina.keyboard.press("Enter")

                await pagina.get_by_role("listbox").filter(has_text="Nacionalidade").get_by_role("combobox").click()
                await pagina.get_by_role("listbox").filter(has_text="Nacionalidade").get_by_role("combobox").fill(row["Nacionalidade:"])
                await pagina.keyboard.press("Enter")

                await pagina.get_by_role("listbox").filter(has_text="Estado Civil").get_by_role("combobox").click()
                await pagina.get_by_role("listbox").filter(has_text="Estado Civil").get_by_role("combobox").fill("Solteiro")
                await pagina.keyboard.press("Enter")
                
                
                await inputFormClick(pagina, "Custo Operacional")
                
                if "." in row["Salário inicial"]:
                    await pagina.keyboard.type(row["Salário inicial"])    
                else:
                    await pagina.keyboard.type(f"{row["Salário inicial"]}00")
                await inputForm(pagina, 'Banco', row["Banco"])
                await inputForm(pagina, 'Agência', row["Número da agência"])
                await inputForm(pagina, 'Nº da conta', handleCaractere(row["Conta e dígito"])["conta"])
                await inputForm(pagina, 'Operador', handleCaractere(row["Conta e dígito"])["digito"])
                if pd.notna(row["Data de início na empresa"]):
                    await inputFormClick(pagina, 'Data de admissão')
                    await pagina.keyboard.type(handleData(row["Data de início na empresa"]))
               # await inputForm(pagina, 'Data de admissão', handleData(row["Data de início na empresa"]))
                
                await pagina.get_by_role("link", name="Endereço").click()
                
                await inputFormClick(pagina, "CEP")
                await pagina.keyboard.type(row["CEP"])
                await asyncio.sleep(2)
               # await inputForm(pagina, 'CEP', row["CEP"])
                await inputForm(pagina, 'Logradouro', row["Logradouro"])
                await pagina.locator('input[formcontrolname="Endereco.Numero"]').fill(row["Número do endereço:"])
                await inputForm(pagina, 'Bairro', row["Bairro"])
                await inputForm(pagina, 'Cidade', row["Cidade"])
                
                await pagina.get_by_role("link", name="Dispositivos de acesso").click()
                
                await pagina.get_by_role("button", name="Adicionar").click()
                await pagina.get_by_role("textbox").fill("12345")
                await pagina.get_by_role("dialog").get_by_role("button", name="Adicionar").click()
                await pagina.get_by_role("button", name="Incluir").click()

            else:
                print(f"{row["Nome Completo:"]} cadastrado atualizar")
            
        await navegador.close()

asyncio.run(main())
