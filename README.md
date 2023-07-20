# PDF from webpage and XLSX from table (dynamic webpage)
Designed to work with custom ApacheSuperset charts.

# JS part

<div style='color:#adf'>
<h3>NOTE</h3> 
Contains custom style fixes and some bug fixes for original packages.
</div>
<br>

## [PDF](./js/html2pdf/html2pdf.js)

Allows you to get the generated pdf(A4 landscape format) from a page with expanded content.

### Run script
```js
getPdf();
```
Returns: binary file object.

## [XLSX](./js/html2xlsx/html2xlsx.js)

Allows you to get a styled xlsx file from a table.

### Run script
```js
saveTable2XLSX($(".pvtTable")[0], "Table title");
```

Returns: binary file object.

## JS requirements:
<div style="color:#adf">
<h3>NOTE</h3> 
These modules are expected to be imported into the page with requirements:

 - PDF:
   - [jQuery](https://cdnjs.com/libraries/jquery)
   - [html2pdf](https://cdnjs.com/libraries/html2pdf.js/0.8.0)
 - XLSX
   - [jQuery](https://cdnjs.com/libraries/jquery)
   - [xlsx-js-style](https://www.jsdelivr.com/package/npm/xlsx-js-style)
</div>

## Python module *selenium_js_executor*
Module contains decorator *exec_js_wrapper* that makes it easier to work with selenium by returning a prepared and configured webdriver.
### Decorator params:
 - **callable**, - decorated function. 
 - **default_driver**:AnyDriver = Firefox, - target *webdriver* for work.
 - **driver_opt**:Dict[str, Any] = {}, - addition options for *webdriver*.
 - **headless**:bool = True, - separated *headless* *webdriver* option. 
 - **window_size**:Tuple[int, int] = (1850, 1132) - separated *window_size* *webdriver* option.

### Decorated params:
 - **url**: str - url of target page.
 - **wait_until**: Optional[*ExceptedCondition*] = None - optional *ExceptedCondition* callable object from *selenium.webdriver.support* when we need wait some event before processing.
 - **wait_timeout**:int = 600 - timeout for condition waiting.
 - **\*args**
 - **\*\*kwargs**


### Example:
#### Using Decorator
```py
# Using default decorator params
@exec_js_wrapper
def bpdf_from_url(driver:'WebDriver', *args, **kwargs) -> bytes:
    """Returns binary pdf of page."""
    script = 'return getPdf();'
    result = base64.b64decode(
        driver.execute_script(script)
    )
    return result
```
The function to be decorated must have the driver as input parameters to work directly with selenium.webdriver.

#### Using Decorated function:
```py
url = "example.com"
wait_condition = selenium_js_executor.expected_conditions.invisibility_of_element_located(
    (selenium_js_executor.By.XPATH, '//*[@id="load"]')
)
timeout = 600

pdf_data_1 = bpdf_from_url(url, wait_condition, timeout)
open('result_1.pdf', 'wb').write(pdf_data_1)
```
