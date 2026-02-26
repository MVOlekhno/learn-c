# VBA_UpdateRevision v26 — конспект и план доработок

Дата разбора: 2026-02-26.

## Что делает макрос
Макрос автоматизирует переход ревизии комплекта РД `C02 → C03`: чистит артефакты, переименовывает файлы/папки, обновляет поля в DOCX/DWG и экспортирует PDF.

## Текущая архитектура
- **Stage 1A**: удаление `*.pdf` и `*.bak` (кроме `ARCHIVE/`, `PDF_ASSEMBLY/`).
- **Stage 1B**: рекурсивное переименование `_C02 → _C03` с нормализацией кириллической/латинской `C`.
- **Stage 2**: обход папок с `docx/dwg`, обработка через Word/AutoCAD COM.

## Подтверждённые риски
1. **S0001 inner folder bug**: внутренняя папка S0001 не должна получать суффикс `_C03`.
2. **Overmatch wildcard дат**: шаблон `NN.NN.24` может случайно задеть не-дату.
3. **Неполная уборка Stage 1A**: остаются `~$*.docx`, `*.dwl`, `*.dwl2`, `hardcopy.log`, `acad.err`.
4. **Кириллическая С в COL_OLD_REV**: при чтении из Excel нужна `NormalizeRev(...)` сразу.

## Приоритетный план (по P0→P3)

### P0
- Починить `RenameFoldersRecursive`: переименовывать только папки, где реально есть `"_" & oldRev` или `"_" & cyrOldRev`.
- Прогнать полную верификацию v26 на комплекте 02UXT по регресс-чеклисту.

### P1
- Добавить удаление `~$*`, `.dwl`, `.dwl2` в `DeleteFilesRecursive`.
- В `UpdateRevision()` читать oldRev через `NormalizeRev(CStr(...))`.

### P2
- Добавить проверку живости и рестарт AutoCAD COM (по аналогии с Word/B-005).
- Ограничить wildcard-замены дат в DOCX проверкой валидности даты перед replace.

### P3
- Добавить поддержку `.doc` в Stage 2.

## Рекомендованные изменения в коде

### 1) Stage 1B: безопасное переименование папок
```vba
If InStr(folderName, "_" & oldRev) > 0 Or InStr(folderName, "_" & cyrOldRev) > 0 Then
    newFolderName = Replace(folderName, "_" & oldRev, "_" & newRev)
    newFolderName = Replace(newFolderName, "_" & cyrOldRev, "_" & newRev)
    If newFolderName <> folderName Then
        fso.MoveFolder currentPath, parentPath & "\" & newFolderName
    End If
End If
```

### 2) Stage 1A: расширение удаления мусора
```vba
If Left(fName, 2) = "~$" Then
    fso.DeleteFile filePath
End If

Select Case LCase(fso.GetExtensionName(filePath))
    Case "pdf", "bak", "dwl", "dwl2"
        fso.DeleteFile filePath
End Select
```

### 3) NormalizeRev сразу при чтении Excel
```vba
oldRev = NormalizeRev(CStr(ws.Cells(r, COL_OLD_REV).Value))
```

### 4) Защита от overmatch в wildcard-датах
- После `Find` проверять найденную подстроку функцией валидации.
- Менять только подтверждённые `DD.MM.YY` и `DD.MM.YYYY`.

## Критерии готовности
- Stage 1 не ломает структуру S0001.
- Все поля DOCX/DWG из чеклиста обновляются.
- Нет зависаний Word/AutoCAD на длинных прогонах.
- Лог `UPDATE_LOG.txt` содержит диагностику по Excel и по каждому файлу.

## Что дальше (практический next step)

Так как в текущем репозитории нет исходника `VBA_UpdateRevision.bas`/`.xlsm`, следующий эффективный шаг — выполнить работу в два спринта.

### Спринт A (сразу)
1. Добавить в репозиторий исходник макроса (`.bas`) и минимальный тестовый комплект папок `C02`.
2. Внести P0-правку в `RenameFoldersRecursive`:
   - переименовывать **только** если имя папки содержит `"_" & oldRev` или `"_" & cyrOldRev`;
   - не трогать inner S0001 без ревизии в имени.
3. Прогнать ручной smoke-тест Stage 1 на копии комплекта и сверить дерево папок.

### Спринт B (после P0)
1. Внести P1-правки:
   - очистка `~$*`, `.dwl`, `.dwl2`, при необходимости `hardcopy.log`, `acad.err`;
   - `oldRev = NormalizeRev(...)` при чтении из Excel.
2. Внести P2-повышения надёжности:
   - health-check/рестарт AutoCAD COM при ошибках открытия;
   - валидация найденной wildcard-строки как даты перед заменой.
3. Прогнать регресс по чек-листу (`PERM`, `UL`, `CAA`, `DWG`) и сохранить артефакты:
   - `UPDATE_LOG.txt`;
   - список файлов до/после;
   - 2-3 контрольных PDF.

## Что нужно предоставить для полноценной реализации здесь
- `VBA_UpdateRevision.bas` (или экспорт модулей VBA),
- образец `.xlsm` с колонками D/F/G/H/J/K/L/M/N,
- обезличенный тестовый набор папок `C02` для 02UXT/08UZT.

После добавления этих файлов можно сделать уже **кодовый** PR (не только документацию) с фиксом P0/P1 и проверкой на реальных данных.

## Формат лога для диагностики зависаний DWG (новое)

При каждом запуске создаётся файл `UPDATE_TRACE.log` в корне выбранной основной папки.

Пример строк:

```
2026-02-26 22:10:01 | Run start OLD=C02 NEW=C03
2026-02-26 22:10:45 | DWG BEGIN: AKU...CLD0001_C03.dwg
2026-02-26 22:10:45 | ProcessDwg open: C:\...\AKU...CLD0001_C03.dwg
2026-02-26 22:10:48 | ProcessDwg opened
2026-02-26 22:12:10 | ProcessDwg handler: [ERR -214...: ...]
```

Что отправлять для разбора:
1. последние 80-120 строк `UPDATE_TRACE.log`;
2. последние 80-120 строк `UPDATE_LOG.txt`;
3. имя зависшего файла DWG.
