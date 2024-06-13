
# Подготовка

1. Создаём папку проекта
2. Открываем в ней терминал
3. Создаём виртуальную среду python - 
```bash
   python3 -m venv env
```
4. Активируем виртуальную среду - 
```bash
   source env/bin/activate
```


5.  Скачиваем все необходимые библиотеки
```bash
pip install pdfrw reportlab openpyxl
```

6. Копируем в папку все необходимые нам файлы:
	1. Solicitud de carnet GVA.pdf
	2. injection.js
	3. pdf_filler.py
7. В браузере заходим в chrome://extensions/ и включаем режим разработчика:![[Pasted image 20240613005310.png]]
8.  Даём `pdf_filler.py` разрешение на запуск:
   ```bash
   chmod +x ./pdf_filler.py
```
# Запуск 

1. В том же терминале где у нас виртуальная среда, мы запускаем наш скрипт `pdf_filler.py`
```bash
   python3 ./pdf_filler.py
```
2. Переходим в браузере на сайт - https://carnetfitosanitarios.es/checkout/
3. Заполняем данные
4. На правую кнопку мыши выбираем `Inspect`, либо на Ctrl+Shift+I, либо на F12 и заход  в `Console`.
5. Пробуем вставить наш скрипт из `injection.js` в консоль и получаем ошибку.![[Pasted image 20240613005440.png]]
6.  Пишем в консоли `allow pasting` и вводим после наш код из `injection.js` 
7. Нажимаем `Realizar el pago`
   ![[Pasted image 20240613005639.png]]
8. Открываем созданный `merged.pdf`![[Pasted image 20240613010205.png]]
9. ???
10. Вы великолепны!
