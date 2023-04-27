# mit import werden die Pakete in das Projekt geladen!
import base64
import datetime
from distutils.log import debug
import io
import dash
import pandas as pd
import webbrowser
import os
import numpy as np
import dash_daq as daq

from threading import Timer

from dash.dependencies import Input, Output, State
from dash import dcc, html, dash_table

# vorgelagerte Funktionen die im Programm verwendet werden
# get_sum_from_kdnr summiert den Marktwert für die jeweilige Kundennummer
def get_sum_from_kdnr(kdnr, result):
    # sum_kurs = sum(result[result['Kundennummer']==kdnr]['Marktwert'])
    sum_kurs = sum(result[result['Kundennummer']==kdnr]['Marktwert in EUR'])
    return sum_kurs

# get_depotsalden liest die Excelfile Depotsalden.exe mit dem Sheet CollExp_SecBalance ein und macht einen Dataframe daraus
def get_depotsalden(depotfile):
    depotsalden = pd.read_excel(depotfile, sheet_name='CollExp_SecBalance')
    return depotsalden

# get_depotsalden liest die Excelfile Kontosalden.exe mit dem Sheet CollExp_SecBalance ein und macht einen Dataframe daraus
def get_kontosalden(kontofile):
    kontosalden = pd.read_excel(kontofile, sheet_name='CollExp_AccBalance')
    kontosalden = pd.DataFrame(kontosalden[['Kundennummer', 'Liqui Total (Kundenwährung)']].groupby('Kundennummer').sum()).reset_index()
    return kontosalden

def open_browser():
    if not os.environ.get("WERKZEUG_RUN_MAIN"):
        webbrowser.open_new('http://127.0.0.1:8049/')


def pad_string(s):
    while len(s) < 12:
        s = '0' + s
    return s

# Nur eine angepasstes Stylesheet
external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

# markiert den Start der Applikation
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

# Hier wird das Layout bestimmt
app.layout = html.Div([
    daq.PowerButton(id='our-power-button-1', on=True, color='#FF5E5E', label='Beenden der Applikation', labelPosition='top'),
    html.Div(id='power-button-result-1'),
    # Uploadbutton mit ID und Style
    dcc.Upload(
        id='upload-data-konto',
        children=html.Div([
            'Bitte die Datei Kontosalden.xlsx in das Fenster einfügen oder ',
            html.A('Durchsuchen')
        ]),
        style={
            'width': '100%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
        },
        multiple=False
    ),
    # Initialisierung der verschiedenen Divs die dann in den Callbacks mit Inhalt gefüllt werden (https://dash.plotly.com/dash-html-components/div) 
    # Generell wird einem Div mit dem Objekt Children Inhalt übergeben. Im Endeffekt sind es einfache HTML-Elemente (https://dash.plotly.com/dash-html-components)
    html.Div(id='output-data-upload-konto'),
    html.Div(id='output-data-upload-test'),
    html.Div(id='output-data-upload-from-dropdown'),
    # Store ist ein unsichtbares Element in dem Daten im Browser zwischengespeichert werden können, die daten werden jeweils ins .json-Format übertragen und dann auch wieder daraus geladen (https://dash.plotly.com/dash-core-components/store)
    dcc.Store(id='kontosalden_json'),
    dcc.Store(id='depotsalden_json'),
    dcc.Store(id='result_json'),
    dcc.Store(id='kurs_json'),
    dcc.Store(id='final_json'),
    # das Elemetn Download kann eine Datei zum Download bereitstellen (https://dash.plotly.com/dash-core-components/download)
    dcc.Download(id="download_dataframe_csv"),
    html.Br()
])

# ein mit @ gekennzeichneter Aufruf ist ein Dekorator und bedeutet nichts anderes, als eine vorgelagerte Funktion die im Hintergrund durchgeführt wird. 
# Es gibt immer einen Output (in der Regel HTML-Elemente, zu speichernde Daten oder Initialbedingungen für nachfolgende Callbacks)
# Als Input werden Elemente beeichnet die eine Aktion, also einen Callback auslösen, z.B. Ändern eines Eintrages im Dropdown Menü oder Button-Click 
# Das State Argument gibt Initialgrößen vor, d.h. es wird kein Callback getriggert, aber die hinterlegten Infos stehen zur Verfügunng und können in den Formeln benutzt werden, 
# z.B. jsonfiles die eingeladen werden oder Inhalte eines Eingabefeldes das zur Filterung genutzt wird
@app.callback(
    #Output ist hier das HTML-Div 'output-data-upload-konto' (Zeile 57) und das Store-Element 'kontosalden_json' (Zeile 61)
    Output('output-data-upload-konto', 'children'),
    Output('kontosalden_json', 'data'),
    # Input und State sind Elemte aus dem dcc.Upload
    Input('upload-data-konto', 'contents'),
    State('upload-data-konto', 'filename'),
    State('upload-data-konto', 'last_modified'),
    prevent_initial_call = True,
    )
# Die folgende Funktion update_output bekommt die Variablen kontofile die dem Input Zeile 80 entspricht, name aus State Zeile 81 und date aus State Zeile 82
def update_output(kontofile, name, date): #name und date werden nicht benötigt
    if kontofile is not None: #Kontrolle ob eine File eingeladen wurde, falls nicht (was bei Start der Seite vorliegt) wird Zeile 122 ausgegeben und in dem entsprechenden Div ('output-data-upload-konto') angezeigt
        kontofile_type, kontofile_string = kontofile.split(',')#Verarbeitung der infos aus dem Input konotfile - String wird zerteilt
        decoded = base64.b64decode(kontofile_string)#deokodiert
        kontofile_decoded = io.BytesIO(decoded)#und in das entsprechende Format übernommen
        try: #bezeichnet den Versuch die Funktion durchzuführen, falls Fail, dann springt er zu Zeile 116 und gibt es im Div aus
            df = get_kontosalden(kontofile_decoded) #Die Exceldatei wird eingelesen und in einen Dataframe gewandelt
            kontosalden = df.to_json(date_format='iso', orient='split') #und abgespeichert im Store-Elemente
            children = html.Div([ #falls alles klappt werden die neuen HTML-Elemente geladen 
                html.H5('Die Datei Kontosalden.xlsx wurde erfolgreich eingelesen'), # Confirm dass die Datei erfolgreich eingeladen wurde
                dcc.Upload( #Upload-Element für Die Depotsalden Datei
                    id='upload-data-depot',
                    children=html.Div([
                        'Bitte die Datei Depotsalden.xlsx in das Fenster einfügen oder ',
                        html.A('Durchsuchen')
                        ]), 
                    style={
                        'width': '100%',
                        'height': '60px',
                        'lineHeight': '60px',
                        'borderWidth': '1px',
                        'borderStyle': 'dashed',
                        'borderRadius': '5px',
                        'textAlign': 'center',
                        'margin': '10px'
                        },
                    multiple=False
                ),
                html.Div(id='output-data-upload-depot'), # leerer Container der im nächsten Schritt befüllt wird
                ])
        except:
            children = html.Div([
                html.H5('Die einzulesende Datei hat nicht das entsprechende Format!'),
                ])            
            init_kontosalden = pd.DataFrame(columns = ['Name', 'Articles', 'Improved'])
            kontosalden = init_kontosalden.to_json(date_format='iso', orient='split')
        return children, kontosalden
    else:
        children = html.Div([html.H5('Bitte die Datei Kontosalden.xlsx in das Fenster einfügen!'),])
        return children

# Gleiche vorgehensweise wie oben
@app.callback(
    Output('output-data-upload-depot', 'children'),
    Output('depotsalden_json', 'data'),
    Input('upload-data-depot', 'contents'),
    State('upload-data-depot', 'filename'),
    State('upload-data-depot', 'last_modified'),
    prevent_initial_call = True,
    )
def update_output(depotfile, name, date):
    if depotfile is not None:
        depotfile_type, depotfile_string = depotfile.split(',')
        decoded = base64.b64decode(depotfile_string)
        depotfile_decoded = io.BytesIO(decoded)
        try:
            df = get_depotsalden(depotfile_decoded)
            depotsalden = df.to_json(date_format='iso', orient='split')
            text = 'Die Datei Depotsalden.xlsx wurde erfolgreich eingelesen'
        except:
            text = 'Die einzulesende Datei hat nicht das entsprechende Format!'
            init_depotsalden = pd.DataFrame(columns = ['Name', 'Articles', 'Improved'])
            depotsalden = init_depotsalden.to_json(date_format='iso', orient='split')
        children = html.Div([
            html.H5(text),
            ])
        return children, depotsalden
    else:
        children = html.Div([html.H5('Bitte die Datei Depotsalden.xlsx in das Fenster einfügen!'),])
        return children
# Input hier sind die beiden Tabellen im json-Format, als Output werden die beiden Tabellen zusammengeführt und wieder als Json gespeichert
@app.callback(
    Output('output-data-upload-test', 'children'),
    Output('result_json', 'data'),
    State('kontosalden_json', 'data'),
    Input('depotsalden_json', 'data'),
    prevent_initial_call = True,
    )
def update_output(kontosalden_json, depotsalden_json):
    try:
        dff_konto = pd.read_json(kontosalden_json, orient='split') # einlesen der Files
        dff_depot = pd.read_json(depotsalden_json, orient='split')
        result = dff_depot.merge(dff_konto, on="Kundennummer", how="left") # hier wird ein merge durchgeführt, vergleichbar mit einem SQL-LeftJoin (https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.merge.html)
        result['Portfoliowert'] = [get_sum_from_kdnr(x, result) for x in result['Kundennummer']] # hier wird in einer Schleife über die Spalte 'Kundennummer' iteriert und jeweils die Fuunktion get_sum_from_kdnr ausgeführt 
        #die den Marktwert der einzelnen Titel aufsummiert und in die Spalte Portfoliowert schreibt
        result['Gesamtwert'] = result['Portfoliowert'] + result['Liqui Total (Kundenwährung)'] # Portfoliowert wird mit dem Liquiwert summiert und ergibt den Gesamtwert
        result['Wertpapier_kürzel'] = result['Kürzel']
        result['Kürzel'] = result['WKN'].str.cat(result['Wertpapier_kürzel'], sep =" - ")
        resultstore = result.to_json(date_format='iso', orient='split') # abspeicherung
        result_titel = result['Kürzel'].unique() # hier wird eine Liste mit den Kürzeln erstellt ohne Doppelnennungen (unique() nimmt jedes Element nur einmal in die Liste auf)
        result_titel = np.append(result_titel, 'Wertpapier noch nicht in der Liste') # neuen eintrag an erster Stelle der Liste erstellen
        children = html.Div([
            html.H5('Bitte das gewünschte Kürzel auswählen!'),
            dcc.Dropdown(result_titel, result_titel[-1], id='titel_dropdown'), # die Liste mit den Titeln/Kürzeln wird als Auswahlmöglichkeiten in dem Dropdownfenster initialisiert welches auch mit in das Div übergeben wird
            ])
    except:
        children = html.Div([
            html.H5('Irgendetwas hat hier nicht funktioniert!'),
            ])
        init_resultstore = pd.DataFrame(columns = ['Name', 'Articles', 'Improved'])
        resultstore = init_resultstore.to_json(date_format='iso', orient='split')
    return children, resultstore

# die Auswahl des Kürzels aus der Dropdownliste und der zusammengefügte (gemerged) Dataframe sind der Input für die folgenden Funktion die den nach Kürzel gefilterten Dataframe herausgibt 
@app.callback(
    Output('output-data-upload-from-dropdown', 'children'),
    Output('kurs_json', 'data'),
    State('result_json', 'data'),
    Input('titel_dropdown', 'value'),
    prevent_initial_call = True,
    )
def update_output(result_json, titel):
    if titel == 'Wertpapier noch nicht in der Liste':
        dff_result = pd.read_json(result_json, orient='split') # einlesen zusammengefügter Dataframe
        dff_result['Gesamtwert'] = dff_result['Portfoliowert'] + dff_result['Liqui Total (Kundenwährung)']
        dff_result = dff_result.drop_duplicates(subset=['Kundennummer'])
        dff_result['Anzahl / Nominal'] = 0
        dff_result['Aktueller Kurs'] = 0
        dff_result['Marktwert'] = 0
        dff_result['Marktwert in EUR'] = 0
        dff_result['Kürzel'] = 'Wertpapier noch nicht in der Liste'
        dff_result['WKN'] = 'WKN'
        dff_result['Wertpapier'] = 'Wertpapier'
        dff_result['Geschäftsbereich'] = 'Geschäftsbereich' 
        df_titel = dff_result[dff_result['Kürzel'] == titel] # Filtern des Dataframes mit dem Kürzel titel aus der Dropdownauswahl        
        resultstore = df_titel.to_json(date_format='iso', orient='split') # speichern im json
        kurs = df_titel['Aktueller Kurs'].unique()[0] # den Kurs aus der Liste auswählen als Initialbedingung für das Eingabefeld Zeile 193
        children = html.Div([
                html.H5('Quote und den aktuellen Kurs auswählen. Mit Bestätigung wird die Übersicht erstellt!'), # HTML-Element 
                dcc.Input(id="input_wkn", type="text", debounce=True, placeholder='WKN eintragen'),
                # html.H5(id="input_wkn", value='Text')
                dcc.Input(id="input_prozent", type="number", debounce=True, value=2), # Eingabeelement für den gewünschten Prozentsatz
                dcc.Input(id="input_value", type="number", debounce=True, value=kurs), # Eingabeelement für den aktuellen Kurs, Initial mit dem Kurs aus der Excelfile befüllt
                html.Button('bestätigen', id='submit_values'), # Button zum bestätigen der Eingabe
                html.Div(id='output-data-final') # div das später befüllt wird
                ])
    else:
        dff_result = pd.read_json(result_json, orient='split') # einlesen zusammengefügter Dataframe 
        df_titel = dff_result[dff_result['Kürzel'] == titel] # Filtern des Dataframes mit dem Kürzel titel aus der Dropdownauswahl
        resultstore = df_titel.to_json(date_format='iso', orient='split') # speichern im json
        kurs = df_titel['Aktueller Kurs'].unique()[0] # den Kurs aus der Liste auswählen als Initialbedingung für das Eingabefeld Zeile 193
        children = html.Div([
                html.H5('Quote und den aktuellen Kurs auswählen. Mit Bestätigung wird die Übersicht erstellt!'), # HTML-Element 
                dcc.Input(id="input_wkn", type="text", debounce=True, style= {'display': 'none'}),
                dcc.Input(id="input_prozent", type="number", debounce=True, value=2), # Eingabeelement für den gewünschten Prozentsatz
                dcc.Input(id="input_value", type="number", debounce=True, value=kurs), # Eingabeelement für den aktuellen Kurs, Initial mit dem Kurs aus der Excelfile befüllt
                html.Button('bestätigen', id='submit_values'), # Button zum bestätigen der Eingabe
                html.Div(id='output-data-final') # div das später befüllt wird
                ])
    return children, resultstore

# Input bzw. State sind die Eingaben aus den Feldern, also Prozentsatz und aktueller Kurs, die Json-File mit den gefiltertetn Werten
# Als Auslöser wird der Klick auf den Button gewählt, hier wird der Callback gestartet
@app.callback(
    Output('output-data-final', 'children'),
    Output('final_json', 'data'),
    Input('submit_values', 'n_clicks'),
    State('kurs_json', 'data'),
    State('input_prozent', 'value'),
    State('input_value', 'value'),
    State('input_wkn', 'value'),
    prevent_initial_call = True,
    )
def update_output(n_clicks, kurs_json, input_prozent, input_value, input_wkn):
    if n_clicks >= 1 and input_prozent and input_value:
        df_filtered = pd.read_json(kurs_json, orient='split') # einlesen der File aus der Zwischenspeicherung
        if input_wkn:
            df_filtered['Aktueller Kurs'] = input_value
            df_filtered['Marktwert'] = 0
            df_filtered['Marktwert in EUR'] = 0
            df_filtered['Kürzel'] = input_wkn
            df_filtered['WKN'] = input_wkn
        else:
            df_filtered['Aktueller Kurs'] = input_value
        if input_prozent == 0:
            df_filtered['Zielstückzahl'] = 0
        else:
            df_filtered['Zielstückzahl'] = [round(x * (input_prozent/100) / input_value, 0) for x in df_filtered['Gesamtwert']] # Schleife über die Spalte Gesamtwert in der mit dem jeweiligen Kurs und dem gewünschten Prozentsatz die Zielstückzahl berechnet wird
        df_filtered['Trade'] = df_filtered['Zielstückzahl'] - df_filtered['Anzahl / Nominal'] # Abgelich der bereits vorhandenen Stückzahl mit der Zielstückzahl
        df_filtered['aktuelle Quote'] = round((df_filtered['Anzahl / Nominal'] * input_value)/(df_filtered['Gesamtwert']/100),2) # aktuelle Prozentquote der Aktie im Depot
        df_filtered['gewünschte Quote'] = input_prozent # gewünschte Prozentquote im Depot (Vorgabe durch eingabefeld)
        df_filtered_table = df_filtered[['Kundennummer', 'Nachname', 'Vorname', 'Depot', 'Gesamtwert', 'Anzahl / Nominal', 'aktuelle Quote', 'gewünschte Quote', 'Zielstückzahl', 'Trade']] # Auswahl der Spalten die im Browser angezeigt werden sollen
        resultstore = df_filtered.to_json(date_format='iso', orient='split')
        children = html.Div([
                html.H5('Mit dem Button Download wird die .csv-Datei heruntergeladen!'),
                html.Button('.csv-Datei herunterladen', id='download'), # Button zum Download der Datei
                html.Br(),
                html.Div(id='output-success'),
                dash_table.DataTable(df_filtered_table.to_dict('records'),[{"name": i, "id": i} for i in df_filtered_table.columns], id='tbl'), # visualisierung des Dataframes zur Anzeige und Prüfung im Browser 
                ])
        return children, resultstore
    else:
        children = html.Div([
                html.H5('Etwas hat nicht funktioniert!')
                ])
        return children

# mit Klick auf den Button wird aus dem übergebenen Dataframe eine ';'-getrennte Datei erstellt und über das Download-Element (Zeile 253) augegeben
@app.callback(
    Output('output-success', 'children'),
    Output("download_dataframe_csv", "data"),
    Input('download', 'n_clicks'),
    State('final_json', 'data'),
    prevent_initial_call = True,
    )
def update_output(n_clicks, final_json):
    if n_clicks >= 1:
        dff_filtered = pd.read_json(final_json, orient='split').reset_index() #einlesen und zurücksetzen des Indexes
        df_export = dff_filtered[['Depot', 'Trade']] # reduzieren auf Depot und Trade (zu tradende Stückzahl)
        df_export['depotstring'] = [str(x) for x in df_export['Depot']] # umwandelung der Depotnummer in einen String
        # Note: 26.04.2023 - neue Spalte mit '0' vor der Depotnummer auffüllen bis 12 Zeichen erreicht sind, auch als String
        df_export['depotstring_zeros'] = [pad_string(x) for x in df_export['depotstring']] 
        # df_export['Trade_betrag'] = round(df_export['Trade'].abs(),0) # Betrag der Zielstückzahl um negative Elemente auszuschließen
        df_export['Trade_betrag'] = round(df_export['Trade'],0) # Betrag der Zielstückzahl um negative Elemente auszuschließen
        df_export.Trade_betrag = df_export.Trade_betrag.astype(int) # umwandlung in Int um aus 1.00 eine 1 zu machen
        df_export['Trade_betrag'] = df_export['Trade_betrag'].astype(str) # wieder speicherung als String
        df_export[''] = ''
        df_csv = df_export[['depotstring_zeros', 'depotstring', '' ,'Trade_betrag']] # erstellung des finalen Dataframes der die zu exportierenden Größen enthält
        df_csv['AUFTRAGSART:STUECKE'] = df_csv.apply(lambda x: ';'.join(x.dropna().values), axis=1) # neue Spalte erstellen die die drei anderen spalten mit einem ';' getrennt generiert
        # df_excel = df_csv['AUFTRAGSART:STUECKE'].apply('="{}"'.format)
        name = str(dff_filtered['Kürzel'][0]) # name der Datei aus der Spalte der Kürzel extrahieren
        # csv_string = name + '.csv' # Dateiname erstellen
        excel_string = name + '.xlsx' # Dateiname erstellen
        # download = dcc.send_data_frame(df_csv['AUFTRAGSART:STUECKE'].to_csv, csv_string, index = False, header=True) # download-Element initialisieren 
        dff_filtered['AUFTRAGSART:STUECKE'] = df_csv['AUFTRAGSART:STUECKE'].apply('="{}"'.format)
        dff_filtered = dff_filtered[['Nachname', 'Vorname', 'Depot', 'Trade', 'AUFTRAGSART:STUECKE']]
        download = dcc.send_data_frame(dff_filtered.to_excel, excel_string, sheet_name = "Export", index = False, header=True) # download-Element initialisieren 
        children = html.Div([
                html.H5('Die Datei wurde erfolgreich erstellt!', style={'color': 'green', 'fontSize': 14}), # Ausgabe Erfolgsmedlung
                ])
        return children, download # Auslieferung der finalen Datei
    else:
        children = html.Div([
                html.H5('Etwas hat nicht funktioniert!')
                ])
        return children


@app.callback(
    Output('power-button-result-1', 'children'),
    Input('our-power-button-1', 'on'),
    prevent_initial_call = True,
)
def update_output(on):
    os.system("taskkill /T /F /im app.exe")
    return 'Die Applikation wurde erfogreich beendet!'

if __name__ == '__main__':
    Timer(1, open_browser).start()
    app.run_server(port=8049)
    # app.run_server(port=8050)
