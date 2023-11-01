import asyncio
import requests
import openpyxl
import pyppeteer
import json
import sys
import logging
from bs4 import BeautifulSoup
from bs4.element import Tag, NavigableString
from pyppeteer import launch
from telebot import TeleBot
from telebot.types import Message

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] [%(name)s] [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

TOKEN = "6868186294:AAG2cIa6nyijkUDuLBIY8RlWanjI9-4_N1E"

bot = TeleBot(TOKEN)


async def main(url, sheet):
    logger.info('requests started!')
    resp = requests.get(url)
    soup = BeautifulSoup(resp.text, 'lxml')
    tbody = soup.find('tbody', id='grid_tab')
    url_fillers = []
    DATA = []
    for i in tbody.contents:
        if isinstance(i, Tag) and i.name == 'tr':
            url_fillers.append((i.attrs['idt'], i.attrs['idf'], i.attrs['idb']))
            row_data = []
            for j in i.contents:
                if isinstance(j, Tag) and j.name == 'td':
                    if j.text.strip() == '':
                        # if ['no14001 _ae'] in j.attrs['class']:
                            # print(j)
                        if 'class' not in j.attrs.keys() or '_ae' in j.attrs['class']:
                            continue
                            
                    row_data.append(j.text.strip())
            DATA.append(row_data)
    logger.info(f'{len(DATA)} rows found in {url}')
    
    titles = ['Продукция', 'Размер', 'Марка', 'Длина',
              'Регион', 'Цена за 1 т', 'Цена, руб до 100м',
              'Цена, руб от 100 до 360м', 'Цена, руб от 360м']
    
    titles.append('Количество')
    DATA.insert(0, titles)
    

    brow = await launch(headless=True)
    pages = await brow.pages()
    page = pages[0]
    k = 0
    for k, i in enumerate(url_fillers, start=1):
        id, idf, idb = i
        logger.info(f'{k} - rows in the queue')
        url = f"https://mc.ru/pages/blocks/add_basket.asp/id/{id}/idf/{idf}/idb/{idb}/action/add"
        await page.goto(url)
        inp = '#tonns'
        await page.click(inp)
        await page.keyboard.type('99999999')
        btn = '.grayBtn'
        await page.click(btn)
        t = await _evaluate(page)
        sys.stdout.write('\x1b[1A')
        if t is None:
            logger.error(f'{k}-row is {t} in {url}')
        else:
            title = await page.evaluate('(element) => element.textContent', t)
            title = title.strip()
            if title == 'Указанного количества нет на складах':
                title = 0
            else:
                title = ' '.join(title.split()[-2:])
            DATA[k].append(title)
        k += 1

    await brow.close()

    for d in DATA:
        sheet.append(d)
    logger.info(f'{k} rows succesfull parsed in this url: {url}')


async def _evaluate(page, retries=0):
    if retries > 10:
        txt = await page.querySelector('.error')
        # return await page.evaluate('document.body.innerHTML')
        return txt
    else:
        try:
            txt = await page.querySelector('.error')
            if txt is None:
                await page.waitFor(500)
                txt = await page.querySelector('.error')
            return txt
            # return await page.evaluate('document.body.innerHTML')
        except pyppeteer.errors.NetworkError:
             await page.waitFor(500)
             return await _evaluate(page, retries+1)


@bot.message_handler(commands=['start'])
def start_bot(message: Message):
    chat_id = message.chat.id
    bot.send_message(chat_id, 'Привествие 👋')
    bot.send_message(chat_id, 'Используйте команду /file, чтобы получить файл')


@bot.message_handler(commands=['file'])
def bot_send_file(message: Message):
    chat_id = message.chat.id
    mess = bot.reply_to(message, 'Сбор данных...')

    workbook = openpyxl.Workbook()
    URLS = []
    titles = []

    with open('template_.json', 'r', encoding='utf-8') as file:
        dc = json.load(file)
        URLS = dc['urls']
        titles = dc['titles']
    
    for i in range(len(URLS)):
        sheet = workbook.create_sheet(titles[i], i+1)
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(main(URLS[i], sheet))
        loop.close()
        workbook.save('mc_ru_data.xlsx')
    bot.delete_message(chat_id, mess.message_id)
    with open('mc_ru_data.xlsx', 'rb', encoding='utf-8') as file:
        bot.send_document(chat_id, file)


if __name__ == '__main__':
    bot.infinity_polling()
