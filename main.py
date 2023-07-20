import base64
import os

import selenium_js_executor


TEST_URL_PDF_1 = '{url}/superset/dashboard/1/?permalink_key=3DbQDPNQJ1a&standalone=2&show_filters=0'
TEST_URL_PDF_2 = '{url}/superset/dashboard/1/?permalink_key=PLVpNBNpvxK&standalone=2&show_filters=0'

TEST_URL_XLSX_1 = '{url}/superset/dashboard/2/?permalink_key=xNPqe0Dp65M&standalone=2&show_filters=0'
TEST_URL_XLSX_2 = '{url}/superset/dashboard/4/?permalink_key=xNPqe0Dp65M&standalone=2&show_filters=0'


@selenium_js_executor.exec_js_wrapper
def bxlsx_from_url(driver) -> bytes:

    script = 'return saveTable2XLSX(document.getElementsByClassName("pvtTable")[0], "test_title");'

    result = base64.b64decode(
        driver.execute_script(script)
    )
    return result


@selenium_js_executor.exec_js_wrapper
def bpdf_from_url(driver) -> bytes:
    script = 'return getPdf();'
    result = base64.b64decode(
        driver.execute_script(script)
    )
    return result


if __name__ == '__main__':
    url = os.getenv('TARGET_URL')

    wait_condition = selenium_js_executor.expected_conditions.invisibility_of_element_located(
        (selenium_js_executor.By.XPATH, '//*[@id="load"]')
    )
    timeout = 600

    pdf_data_1 = bpdf_from_url(TEST_URL_PDF_1.format(url=url), wait_condition, timeout)
    open('result_1.pdf', 'wb').write(pdf_data_1)
    pdf_data_2 = bpdf_from_url(TEST_URL_PDF_2.format(url=url), wait_condition, timeout)
    open('result_2.pdf', 'wb').write(pdf_data_2)

    xlsx_data_1 = bxlsx_from_url(TEST_URL_XLSX_1.format(url=url), wait_condition, timeout)
    open('result_1.xlsx', 'wb').write(xlsx_data_1)
    xlsx_data_2 = bxlsx_from_url(TEST_URL_XLSX_2.format(url=url), wait_condition, timeout)
    open('result_2.xlsx', 'wb').write(xlsx_data_2)
