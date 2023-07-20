from typing import Any, Callable, Dict, Optional, Tuple, Union

from time import sleep

from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver import Firefox, Chrome, Edge, Safari, FirefoxOptions, ChromeOptions, EdgeOptions
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait


# conditions from selenium.webdriver.support.expected_conditions
ExceptedCondition = Callable[[Union[WebElement, Tuple[str, str]]], Callable[[expected_conditions.AnyDriver], Union[WebElement, bool]]]

def _get_default_options(default_driver:expected_conditions.AnyDriver) -> Union[FirefoxOptions, ChromeOptions, EdgeOptions]:
    """Returns"""
    if default_driver == Firefox:
        return FirefoxOptions()
    elif default_driver == Chrome:
        return ChromeOptions()
    elif default_driver == Edge:
        return EdgeOptions()
    elif default_driver == Safari:
        raise Exception("Safari options are not avaliable now!")
    else:
        raise Exception(f"Error: Unknown webdriver! {type(default_driver)}")


def exec_js_wrapper(
    callable,
    default_driver:expected_conditions.AnyDriver = Firefox,
    driver_opt:Dict[str, Any] = {},
    headless:bool = True,
    window_size:Tuple[int, int] = (1850, 1132)
    ) -> Callable[[str, Optional[ExceptedCondition], int], str]:

    options = _get_default_options(default_driver)
    for key, value in driver_opt.items():
        options[key] = value
    options.headless = headless

    def from_url(
        url: str,
        wait_until: Optional[ExceptedCondition] = None,
        wait_timeout:int = 600,
        *args,
        **kwargs
    ) -> str:
        driver:WebDriver = default_driver(options= options)
        driver.get(url)
        driver.set_window_size(*window_size)

        try:
            WebDriverWait(driver, wait_timeout).until(wait_until)
        except TimeoutException:
            print('TimeoutError!')
        except AttributeError:
            pass

        sleep(1)
        result = callable(driver, *args, **kwargs)
        sleep(2)
        driver.close()
        return result

    return from_url
