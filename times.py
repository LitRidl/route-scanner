from sys import argv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from random import randint as ri


home = 'Москва, Лефортово, ЖК Символ, к. 3'
locations = [
    'БЦ Метрополис', 
    'БЦ Алкон',
    'БЦ Водный',
    'БЦ Marina Business Park',
    'БЦ Большевик',
    'МФК Gardenmir',
    'бизнес-квартал Голутвинская слобода',
    'БЦ Монарх',
    'БЦ Виктори Плаза',
    'БЦ Лялин 19 к1',
    'электротеатр Станиславского',
    'БЦ Оазис',
    'БЦ Арена Парк',
    'БЦ Лопухинский 3 с2',
    'БЦ Нарышкинская 5 с1',
    'БЦ Ленинградское 39 с2',
    'ЖК Триумф Палас',
    'деловой центр Golden Gate',
    'БЦ Auroom',
    'бизнес-квартал Красная Роза',
    'БЦ Lighthouse',
    'Бобров переулок, 4, стр. 1',
    'БЦ Порт Плаза',
    'БЦ Бригантина Холл',
    'БЦ Арма',
    'Бизнес центр Neo Geo'
    'БЦ Белая площадь',
    'БЦ Ситидел',
    'БЦ White Stone'
    'апарт-квартал Tribeca Apartments', 
    'МФК башня Империя'
]


options = webdriver.ChromeOptions()
options.binary_location = "/Applications/Chrome.app/Contents/MacOS/Google Chrome"
chrome_driver_binary = "/usr/local/bin/chromedriver"
driver = webdriver.Chrome(chrome_driver_binary, chrome_options=options)

driver.get('https://yandex.ru/maps/213/moscow/?mode=routes&rtext=&rtt=auto')

src_input, dst_input = driver.find_elements_by_class_name('input_medium__control')

from_home = {}
for loc in locations:
    ActionChains(driver).pause(ri(3, 6)).move_to_element(src_input).send_keys(home).send_keys(Keys.RETURN).perform()
    ActionChains(driver).move_to_element(dst_input).send_keys('Москва, ' + loc).send_keys(Keys.RETURN).pause(ri(2, 4)).perform()

    routes = driver.find_elements_by_class_name('driving-route-view__route-title-primary')
    times = [r.text.replace('мин', '').strip() for r in routes]

    from_home[loc] = times

    ActionChains(driver).move_to_element(driver.find_element_by_class_name('route-form-view__reset')).click().pause(1).perform()

print(from_home)
driver.close()

if __name__ == '__main__':
    pass
    # locations = ['Moscow', 'Saint-Petersburg', 'Novosibirsk']
    # locations = argv[1:]
    # if len(locations) == 0:
    #     print("Пример использования:\n\t{0} moscow saint-petersburg novosibirsk".format(argv[0]))

    # print(get_times(home, other))    
    # try:
    #     workbook = load_workbook(weather_filename)
    # except:
    #     workbook = Workbook()
    #     workbook.active.append(['Местоположение', 'Дата', 'Реальная температура', 'Ощущаемая температура', 'Влажность', 'Метеоусловия'])

    # sheet = workbook.active
    # for location in locations:
    #     w = get_weather(location)
    #     if w is None:
    #         print('Ошибка: местоположение "{0} не найдено"'.format(location))
    #         continue
    #     sheet.append([location, w['day'], w['t_fact'], w['t_feels_like'], w['humidity'], w['condition']])
        

    # workbook.save(weather_filename)
