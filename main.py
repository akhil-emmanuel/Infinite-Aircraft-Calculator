import openpyxl
import PySimpleGUI as sg
from openpyxl import Workbook, load_workbook


book = load_workbook('IF Aircraft Data.xlsx')

def round_nrst_10(number):
    return round(number / 10) * 10

def InputLoadtoDataRow(acLoad):
    acLoad = int(acLoad)
    if acLoad == 23 or acLoad == 24 or acLoad == 25 or acLoad == 26 or acLoad == 27:
        row = 10
    elif acLoad == 73 or acLoad == 74 or acLoad == 75 or acLoad == 76 or acLoad == 77:
        row = 4
    else:
        acLoad = round_nrst_10(acLoad)
        if acLoad == 0 or acLoad == 10:
            row = 11
        elif acLoad == 20:
            row = 10
        elif acLoad == 30:
            row = 9
        elif acLoad == 40:
            row = 8
        elif acLoad == 50:
            row = 7
        elif acLoad == 60:
            row = 6
        elif acLoad == 70:
            row = 5
        elif acLoad == 80:
            row = 4
        elif acLoad == 90:
            row = 3
        elif acLoad == 100:
            row = 2
    #print(row)
    return str(row)

def getDepatureData(row):
    DepPower = sheet['B' + row].value
    DepFlaps = sheet['C' + row].value
    DepRotate = sheet['D' + row].value
    DepAirBy = sheet['E' + row].value
    return DepPower, DepFlaps, DepRotate, DepAirBy
def getArrivalData(row):
    LdgFlaps = sheet['F'+ str(row)].value
    LdgApprSpd  = sheet['G'+ str(row)].value
    LdgFlareSpd = sheet['H'+ str(row)].value
    FlapsSpd = sheet['A14'].value
    return LdgFlaps, LdgApprSpd, LdgFlareSpd, FlapsSpd

def getFuelBurnData(row):
    Even = sheet['I'+ str(row)].value
    Odd = sheet['J'+ str(row)].value
    MedBurn = sheet['K'+ str(row)].value
    RecWest = sheet['E14'].value
    RecEast = sheet['F14'].value
    return Even, Odd, MedBurn, RecWest, RecEast

def getOtherData(row):
    Ceiling = sheet['B14'].value
    Cruise = sheet['C14'].value
    MMO = sheet['D14'].value
    Range = sheet['A13'].value
    return Ceiling, Cruise, MMO, Range

def CL350(oat):
    if oat == -30:
        return "80% = 81.4% N1"
    elif oat == -25:
        return "81% = 82.3% N1"
    elif oat == -20:
        return "82% = 83.2% N1"
    elif oat == -15:
        return "83% = 84.0% N1"
    elif oat == -10:
        return "84% = 84.7% N1"
    elif oat == -5:
        return "85% = 85.5% N1"
    elif oat == 0:
        return "87% = 86.3% N1"
    elif oat == 5:
        return "87% = 87.0% N1"
    elif oat == 10:
        return "88% = 87.8% N1"
    elif oat == 15:
        return "90% = 88.7% N1"
    elif oat == 20:
        return "91% = 89.5% N1"
    elif oat == 25:
        return "92% = 90.3% N1"
    elif oat == 30 or oat == 35:
        return "93% = 91.1% N1"
    elif oat == 40:
        return "90% = 88.7% N1"
    else:
        return "Invalid input"

# Define your custom theme
custom_theme = {
    'BACKGROUND': '#1C1C1D',
    'TEXT': '#FFFFFF',
    'INPUT': '#404040',
    'TEXT_INPUT': '#FFFFFF',
    'SCROLL': '#404040',
    'BUTTON': ('#FFFFFF', '#404040'),
    'PROGRESS': ('#FFFFFF', '#D0D0D0'),
    'BORDER': 1,
    'SLIDER_DEPTH': 0,
    'PROGRESS_DEPTH': 0,
}

# Set the theme
sg.theme_add_new('CustomTheme', custom_theme)
sg.theme('CustomTheme')
#sg.theme('DarkGrey6')

layout = [
    [sg.Image(filename='Title.png')],
    [sg.Frame('Input', layout=[
        [sg.Text('Aircraft Manufacturer:'), sg.Button('Airbus'), sg.Button('Boeing'), sg.Button('Bombardier'),
         sg.Button('Embraer'), sg.Button('McDonnell Douglas')],
        [sg.Text('Aircraft Type:'), sg.DropDown(['            '], key='selected_aircraft')],
        [sg.Text('Aircraft Load:'), sg.InputText(key='load', size=(3, 1)), sg.Text('%')],

        #CL350 ONLY
        [sg.Text("Select OAT (°C):", key='askOAT', visible=False), sg.Slider(range=(-30, 40), default_value=0, orientation="h", key="OATinput", size=(20, 20), resolution=5, visible=False)],

        [sg.Text('Request:'), sg.Button('Departure Data'),  sg.Button('Fuel Burn Data'), sg.Button('Arrival Data'), sg.Button('Other Data')]
    ])],[sg.Frame('Output', layout=[
        [sg.Text(key='output', text_color='grey', enable_events=True )],
    ])],
        [sg.Button("Credits"), sg.Button("Quit")]

]


window = sg.Window("Infinite Aircraft Calculator", layout, font=('Helvetica Neue', 14), size=(640, 550), resizable=True,)



while True:
    event, value = window.read()
    if event == "Quit" or event == sg.WIN_CLOSED:
        break

    elif event == 'Airbus':
        window['Airbus'].update(button_color = ('#404040','#FFFFFF' )); window['Boeing'].update(button_color=('#FFFFFF', '#404040')); window['Bombardier'].update(button_color=('#FFFFFF', '#404040')); window['Embraer'].update(button_color=('#FFFFFF', '#404040')); window['McDonnell Douglas'].update(button_color=('#FFFFFF', '#404040'))
        new_choices = ['A220', 'A318', 'A319', 'A320', 'A321', 'A332', 'A333', 'A339', 'A346', 'A359', 'A388']
        window['selected_aircraft'].update(values=new_choices)

    elif event == 'Boeing':
        window['Boeing'].update(button_color = ('#404040','#FFFFFF' )); window['Airbus'].update(button_color=('#FFFFFF', '#404040')); window['Bombardier'].update(button_color=('#FFFFFF', '#404040')); window['Embraer'].update(button_color=('#FFFFFF', '#404040')); window['McDonnell Douglas'].update(button_color=('#FFFFFF', '#404040'))
        new_choices = ['B712', 'B737', 'B738', 'B739', 'B742', 'B744', 'B749', 'B752', 'B763', 'B772', 'B77L', 'B77W', 'B77F',
               'B788', 'B789', 'B78X']
        window['selected_aircraft'].update(values=new_choices)

    elif event == 'Bombardier':
        window['Bombardier'].update(button_color = ('#404040','#FFFFFF' )); window['Boeing'].update(button_color=('#FFFFFF', '#404040')); window['Airbus'].update(button_color=('#FFFFFF', '#404040')); window['Embraer'].update(button_color=('#FFFFFF', '#404040')); window['McDonnell Douglas'].update(button_color=('#FFFFFF', '#404040'))
        window['Bombardier'].update(button_color=('black', 'white'))
        new_choices = ['CL350', 'CRJ2', 'CRJ7', 'CRJ9', 'CRJX']
        window['selected_aircraft'].update(values=new_choices)

    elif event == 'Embraer':
        window['Embraer'].update(button_color = ('#404040','#FFFFFF' )); window['Boeing'].update(button_color=('#FFFFFF', '#404040')); window['Bombardier'].update(button_color=('#FFFFFF', '#404040')); window['Airbus'].update(button_color=('#FFFFFF', '#404040')); window['McDonnell Douglas'].update(button_color=('#FFFFFF', '#404040'))
        new_choices = ['E175', 'E190']
        window['selected_aircraft'].update(values=new_choices)

    elif event == 'McDonnell Douglas':
        window['McDonnell Douglas'].update(button_color = ('#404040','#FFFFFF' )); window['Boeing'].update(button_color=('#FFFFFF', '#404040')); window['Bombardier'].update(button_color=('#FFFFFF', '#404040')); window['Embraer'].update(button_color=('#FFFFFF', '#404040')); window['Airbus'].update(button_color=('#FFFFFF', '#404040'))
        new_choices = ['DC10', 'DC1F', 'MD11', 'MD1F']
        window['selected_aircraft'].update(values=new_choices)

    elif event == 'Credits':
        window['Credits'].update(button_color = ('#404040','#FFFFFF' )); window['Departure Data'].update(button_color=('#FFFFFF', '#404040')); window['Arrival Data'].update(button_color=('#FFFFFF', '#404040')); window['Fuel Burn Data'].update(button_color=('#FFFFFF', '#404040')); window['Other Data'].update(button_color=('#FFFFFF', '#404040'));
        output_text = '\n'.join([
            f"DeerCrusher: Takeoff and Landing Profile Data for Reworked Aircraft",
            f"Kuba_Jaroszczyk: Takeoff and Landing Profile Data for Older Aircraft",
            f"AndrewWu: Fuel Burn Data and Recommended Flight Profiles",
            f"Jan: Ceiling, Normal Range, Cruise Spd, MMO Spd Data",
            f"darkeyes: The program. Thank you for using.",
            f"\nThe data here isn't perfectly accurate. It is simply \nintended to offer a basic guidance. ",


        ])
        window['output'].update(output_text, text_color='white')

    elif event == 'Departure Data':
        window['Departure Data'].update(button_color = ('#404040','#FFFFFF' )); window['Credits'].update(button_color=('#FFFFFF', '#404040')); window['Arrival Data'].update(button_color=('#FFFFFF', '#404040')); window['Fuel Burn Data'].update(button_color=('#FFFFFF', '#404040')); window['Other Data'].update(button_color=('#FFFFFF', '#404040'));
        try:
            acType = value["selected_aircraft"]
            sheet = book[acType]
            ac_load = value['load']
            row = InputLoadtoDataRow(ac_load)
            DepPower, DepFlaps, DepRotate, DepAirBy = getDepatureData(row)
            if value["selected_aircraft"] == 'CL350':
                window['askOAT'].update(visible=True)
                window['OATinput'].update(visible=True)
                oat = value["OATinput"]
                DepPower = CL350(oat)
                output_text = '\n'.join([f"Flap Setting: {DepFlaps} ", f"\nPower: \n{DepPower} ", f"\nRotate: {DepRotate} ", f"Airborne By: {DepAirBy}"])
            else:
                window['askOAT'].update(visible=False)
                window['OATinput'].update(visible=False)
                output_text = '\n'.join(
                    [f"Flap Setting: {DepFlaps} ", f"\nPower: {DepPower} ", f"\nRotate: {DepRotate} ",
                     f"Airborne By: {DepAirBy}"])
            window['output'].update(output_text, text_color='white')
        except KeyError:
            errorMessage = 'Select Aircraft Type!'
            window['output'].update(errorMessage, text_color='red')
        except ValueError:
            errorMessage = 'Enter Aircraft Load as an Integer!'
            window['output'].update(errorMessage, text_color='red')
        except:
            errorMessage = 'Error, try again'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')

    elif event == 'Arrival Data':
        window['Arrival Data'].update(button_color = ('#404040','#FFFFFF' )); window['Credits'].update(button_color=('#FFFFFF', '#404040')); window['Departure Data'].update(button_color=('#FFFFFF', '#404040')); window['Fuel Burn Data'].update(button_color=('#FFFFFF', '#404040')); window['Other Data'].update(button_color=('#FFFFFF', '#404040'));
        try:
            acType = value["selected_aircraft"]
            sheet = book[acType]
            ac_load = value['load']
            row = InputLoadtoDataRow(ac_load)
            ArrFlaps, ArrApprSpd, ArrFlareSpd, Flaps = getArrivalData(row)
            output_text = '\n'.join([f"Flap Setting: {ArrFlaps} ", f"\nApproach Speed: {ArrApprSpd}", f"Flare Speed: {ArrFlareSpd} ", f"\nFlap Spds: \n{Flaps} "])
            window['output'].update(output_text, text_color='white')
        except KeyError:
            errorMessage = 'Select Aircraft Type!'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')
        except ValueError:
            errorMessage = 'Enter Aircraft Load as an Integer!'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')
        except:
            errorMessage = 'Error, try again'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')

    elif event == 'Fuel Burn Data':
         window['Fuel Burn Data'].update(button_color = ('#404040','#FFFFFF' )); window['Credits'].update(button_color=('#FFFFFF', '#404040')); window['Departure Data'].update(button_color=('#FFFFFF', '#404040')); window['Arrival Data'].update(button_color=('#FFFFFF', '#404040')); window['Other Data'].update(button_color=('#FFFFFF', '#404040'));
         try:
            acType = value["selected_aircraft"]
            sheet = book[acType]
            ac_load = value['load']
            row = InputLoadtoDataRow(ac_load)
            Even, Odd, Med, RecWest, RecEast = getFuelBurnData(row)
            output_text = '\n'.join([f"West/Even Cruise Alt: {Even} ", f"East/Odd Cruise Alt: {Odd} ", f"High Fuel Burn: {Med} ",
                                      f"\n\nRecommend Flight Profile West: {RecWest}", f"\nRecommend Flight Profile East: {RecEast}"])
            window['output'].update(output_text, text_color='white')
         except KeyError:
            errorMessage = 'Select Aircraft Type!'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')
         except ValueError:
             errorMessage = 'Enter Aircraft Load as an Integer!'
             mycolor = 'red'
             window['output'].update(errorMessage, text_color='red')
         except:
             errorMessage = 'Error, try again'
             mycolor = 'red'
             window['output'].update(errorMessage, text_color='red')

    elif event == 'Other Data':
        window['Other Data'].update(button_color = ('#404040','#FFFFFF' )); window['Credits'].update(button_color=('#FFFFFF', '#404040')); window['Departure Data'].update(button_color=('#FFFFFF', '#404040')); window['Arrival Data'].update(button_color=('#FFFFFF', '#404040')); window['Fuel Burn Data'].update(button_color=('#FFFFFF', '#404040'));
        try:
            acType = value["selected_aircraft"]
            sheet = book[acType]
            ac_load = value['load']
            row = InputLoadtoDataRow(ac_load)
            Ceiling, Cruise, MMO, Range = getOtherData(row)
            print(Range)
            output_text = '\n'.join([f"Ceiling: {Ceiling}", f"Normal Range: {Range}", f"\nCruise Spd: {Cruise}", f"MMO Spd: {MMO} \n"])
            window['output'].update(output_text, text_color='white')
        except KeyError:
            errorMessage = 'Select Aircraft Type!'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')
        except ValueError:
            errorMessage = 'Enter Aircraft Load as an Integer!'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')
        except:
            errorMessage = 'Error, try again'
            mycolor = 'red'
            window['output'].update(errorMessage, text_color='red')

window.close()
