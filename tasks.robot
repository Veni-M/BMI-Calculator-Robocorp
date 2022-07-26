*** Settings ***
Documentation       BMI Calculator.

# libraries for browser, excel and tables
Library    RPA.Browser.Selenium       auto_close=${False}
Library    RPA.Excel.Files
Library    String
Library    RPA.Tables


*** Variables ***
${bmiCalculatorUrl}   https://www.iciciprulife.com/tools-and-calculators/bmi-calculator.html
${inputExcelPath}     Input Details.xlsx
${sheetName}          Sheet1
${Gender}       Male
${rowIndex}     2
${counter}      1

*** Keywords ***
Input Details
    Open Workbook    ${inputExcelPath}
    # reading sheet1 of input excel and storing it in the datatable
    ${dataTable}=  Read Worksheet    ${sheetName}   header=True

    # adding "Your BMI" in the input excel
    Set Worksheet Value    1    6    Your BMI
    Save Workbook
    FOR    ${eachRow}    IN    @{dataTable}
        IF    "${eachRow}[Gender]" == "${Gender}"
            Click Element    //label[@for="genderM"]
        ELSE
            Click Element    //label[@for="genderF"]
        END
        Input Text    id=age    ${eachRow}[Age]
        Input Text    id=height_ft    ${eachRow}[Height (Feet)]
        Input Text    id=weight    ${eachRow}[Weight]
        Click Button    Calculate
        Sleep    2
        ${bmi}=   Get Text    //div[@class="result-val js-bmi_txt"]
        Open Workbook    ${inputExcelPath}
        Set Worksheet Value    ${rowIndex}    6    ${bmi}
        Save Workbook
        ${rowIndex}=       Evaluate    ${rowIndex} + ${counter}
    END
     
*** Tasks ***
Open Browser
    Open Available Browser   ${bmiCalculatorUrl}   maximized=True 

Giving Inputs 
    Input Details