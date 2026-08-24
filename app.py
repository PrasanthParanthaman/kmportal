from flask import Flask, render_template, request, jsonify
from flask import session, redirect, url_for
from datetime import timedelta
from datetime import datetime
#import datetime
import os
import process
import pickle
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
from collections import Counter
import re
from threading import Thread
import random
import html
#from talisman import Talisman

unknownResponse = 'I am sorry, but I do not understand.'

# Path to your text file
timestamp_file = "last_update_time.txt"
PREVIOUS_UPDATED_VALUE = None  # Store the previous value globally
        
app = Flask(__name__,static_url_path='/serve')
app.secret_key = 'your_secret_key'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=20)
app.config['SESSION_COOKIE_SECURE'] = True      # Ensures cookies are sent over HTTPS
app.config['SESSION_COOKIE_HTTPONLY'] = True    # Prevents JavaScript access to cookies
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'   # Optional: helps mitigate CSRF
#talisman = Talisman(app)

def getResponse(query):
    with open('trainingData.pickle', 'rb') as handle:
        convoDataStructure=pickle.load(handle)
        try:
            tempResponse = convoDataStructure[query]
            if(isinstance(tempResponse, list)):
                response = random.choice(tempResponse)
            else:
                response = convoDataStructure[query]
        except KeyError:
            response = unknownResponse
        return response

def getCurrentTime():
    return datetime.datetime.now().strftime("[%d-%b-%Y %H:%M:%S.%f]")

def getCurrentDate():
    return datetime.date.today().strftime("%b_%d_%Y")

def Responselogger(folder, typeOfResponse, text, logDate, logTime):
    logFile = open('logs\\{0}Responses\\{1}_Responses_Log_{2}.txt'.format(folder,folder,logDate), 'a')
    logFile.write("{0} {1} => {2}\n".format(typeOfResponse,logTime,text))
    if(typeOfResponse == "BOT"):
        logFile.write("\n")
    logFile.close()

def frequencyAnalysisLogger(userText):
    logFile = open('logs\\frequencyAnalysisData\\FAD.txt', 'a')
    logFile.write("{0} ".format(userText))
    logFile.close()

def chunker(seq, size):
    return (seq[pos:pos + size] for pos in range(0, len(seq), size))

def generateQueryFrequencyAnalysis():
    with PdfPages('static\\generatables\\FAR.pdf') as pdf:   
        dataFile = open("logs\\frequencyAnalysisData\\FAD.txt", "r")
        frequencyData = dataFile.read()
        dataFile.close()
        wordlist = re.sub(r"[,.;@#?!&$]+\ *", " ", frequencyData).split(" ")
        wordlist=sorted(list(filter(None, wordlist)))
        wordFreqTuple=list(Counter(wordlist).items())
        for group in chunker(wordFreqTuple, 56):
            words,occurences=zip(*group)
            plt.rcParams['figure.figsize'] = 15, 10
            plt.barh(words,occurences, align='center', label="Frequency")
            plt.legend()
            plt.ylabel('Words')
            plt.xlabel('Frequency')
            plt.title('Frequency of words in User Queries')
            plt.tight_layout()
            pdf.savefig()
            plt.close()

def queryAdder(query,response):
    try:
        logFile = open("bot_Bin\\conversations\\UserAdded.yml")
        data = logFile.read().strip()
        logFile.close()
    except FileNotFoundError:
        logFile = open('conversations\\UserAdded.yml', 'w')
        data="categories:\n- Terms\nconversations:"
        logFile.write(data)
        logFile.close()   
    logFile = open('conversations\\UserAdded.yml', 'w')
    if data=="":
        logFile.write("categories:\n- Terms\nconversations:")
    logFile.write(data)
    logFile.write("\n")
    logFile.write("- - "+query)
    logFile.write("\n")
    logFile.write("  - "+response)
    logFile.close()
    with open('trainingData.pickle', 'rb') as handle:
        convoDataStructure=pickle.load(handle)
    for currentQuery in " ".join((query.strip()).split()).split("<;>"):
        currentQuery = currentQuery.strip()
        autocorrectTokens = list(convoDataStructure.keys())
        statementMatches = process.extractBests(currentQuery, autocorrectTokens, scorer=process.token_sort_ratio, score_cutoff=1, limit=1)
        actualCutoff = process.token_sort_ratio(currentQuery, statementMatches[0][0])
        if(actualCutoff==100):
            currentQuery = statementMatches[0][0]
            existingResponse = convoDataStructure[currentQuery]
            tempResponseList = []
            if(isinstance(existingResponse, list)):
                for k in existingResponse:
                    tempResponseList.append(k)
            else:
                tempResponseList.append(existingResponse)
            tempResponseList.append(response)
            convoDataStructure[currentQuery]=tempResponseList
        else:
            convoDataStructure[currentQuery]=response       
    with open('trainingData.pickle', 'wb') as handle:
        pickle.dump(convoDataStructure, handle, protocol=pickle.HIGHEST_PROTOCOL)
    return "Query Added Successfully"

def responseProcessor(userText):
    if userText.lower() == "generate far":
        Thread(target=generateQueryFrequencyAnalysis).start()
        response = 'Frequency Analysis Report Generation Started in the Backend - Please wait for a few minutes to the process to be completed<\\n>Use the keyword <font>Download FAR</font> To Download the latest FAR'
    elif userText.lower() == "download far":
        response = '<a class="far_report" href="serve/generatables/FAR.pdf" download>Click Here</a> To Download The Frequency Analysis Report'
    else:
        response = getResponse(userText)
        if(response == unknownResponse):
            with open('trainingData.pickle', 'rb') as handle:
                autocorrectTokens = list(pickle.load(handle).keys())
            statementMatches = process.extractBests(
                userText, autocorrectTokens, scorer=process.token_sort_ratio, score_cutoff=1, limit=4)
            actualCutoff = process.token_sort_ratio(
                userText, statementMatches[0][0])
            if actualCutoff == 100:
                frequencyAnalysisLogger(userText.upper())
                response = getResponse(statementMatches[0][0])
            elif actualCutoff > 80:
                frequencyAnalysisLogger(statementMatches[0][0].upper())
                response = 'Do you mean <font>{0}</font>?<\\n>{1}'.format(statementMatches[0][0],getResponse(statementMatches[0][0]))
            else:
                frequencyAnalysisLogger(userText.upper())
                sussestions = []
                for i in range(len(statementMatches)):
                    sussestions.append(statementMatches[i][0])
                if(len(sussestions) > 0):
                    response = '{0}<\\n>For Google Search: <a class="google_search" href="https://www.google.com/search?q={1}" target="_blank">Click Here</a><\\n>Suggested Close Matches:<br><button class="clickOptions" onClick="clickOptions(this)">{2}</button>'.format(response,userText,'</button><br><button class="clickOptions" onClick="clickOptions(this)">'.join(sussestions))
                else:
                    response = '{0}<\\n>For Google Search: <a class="google_search" href="https://www.google.com/search?q={1}" target="_blank">Click Here</a><\\n>No Close Match Found'.process(response,userText)
        else:
            frequencyAnalysisLogger(userText.upper())
    return response

@app.route('/')
def index():
    session.clear()
    session.permanent = True
    session['user'] = 'User'
    return render_template('index.html')

@app.route('/last-update')
def get_footer_data():
    global PREVIOUS_UPDATED_VALUE
    try:
        # Read the current data from the text file
        with open(timestamp_file, 'r') as file:
            current_data = file.read().strip()

        # Compare with the previous value
        if current_data == PREVIOUS_UPDATED_VALUE:
            response_data = PREVIOUS_UPDATED_VALUE  # Return old value
        else:
            response_data = current_data  # Return new value
            PREVIOUS_UPDATED_VALUE = current_data  # Update previous value
        
        return jsonify({"last_update_time": response_data})
    
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/Adas')
def adas():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Adas.html')

@app.route('/ApplicationManual')
def application_manual():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('ApplicationManual.html')

@app.route('/AskKite')
def ask_kite():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('AskKite.html')

@app.route('/Assessment')
def assessment():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Assessment.html')

@app.route('/AutoKnowAll')
def auto_know_all():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('AutoKnowAll.html')

@app.route('/Automation')
def automation():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Automation.html')

@app.route('/BestPractices')
def best_practices():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('BestPractices.html')

@app.route('/Bms')
def bms():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Bms.html')

@app.route('/Bootcamp')
def bootcamp():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Bootcamp.html')

@app.route('/Casestudies')
def case_studies():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Casestudies.html')

@app.route('/ConnectedCars')
def connected_cars():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('ConnectedCars.html')

@app.route('/Dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Dashboard.html')

@app.route('/DomainVideo')
def domain_video():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('DomainVideo.html')

@app.route('/DriveInDomain')
def drive_in_domain():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('DriveInDomain.html')

@app.route('/ELearnings')
def e_learnings():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('ELearnings.html')

@app.route('/ElectricVehicle')
def electric_vehicle():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('ElectricVehicle.html')

@app.route('/FuelUP')
def fuel_up():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('FuelUP.html')

@app.route('/Guidelines')
def guidelines():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Guidelines.html')

@app.route('/HallOfFame')
def hall_of_fame():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('HallOfFame.html')

@app.route('/HolidayCalendar')
def holiday_calendar():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('HolidayCalendar.html')

@app.route('/HomePage')
def home_page():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('HomePage.html')

@app.route('/Infotainment')
def infotainment():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Infotainment.html')

@app.route('/KeyAccelarator')
def key_accelerator():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('KeyAccelarator.html')

@app.route('/KiteArea')
def kite_area():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('KiteArea.html')

@app.route('/Knowyourdomain')
def know_your_domain():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Kyd.html')

@app.route('/know_your_templates') 
def know_your_templates(): 
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('KnowYourTemplates.html') 

@app.route('/know_your_acronyms') 
def know_your_acronyms():
    if 'user' not in session:
        return redirect(url_for('index'))    
    return render_template('KnowYourAcronyms.html')

@app.route('/KT')
def kt():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('KT.html')

@app.route('/Kya')
def kya():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Kya.html')

@app.route('/Kyc')
def kyc():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Kyc.html')
    
@app.route('/Kyt')
def kyt():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Kyt.html')    

@app.route('/Logout')
def logout():
    session.clear()
    return render_template('Logout.html')

@app.route('/NewsLetter')
def newsletter():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('NewsLetter.html')

@app.route('/OrganizationChart')
def organization_chart():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('OrganizationChart.html')

@app.route('/ProjectTimeline')
def project_timeline():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('ProjectTimeline.html')
    
@app.route('/PromptBook')
def prompt_book():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('PromptBook.html')    

@app.route('/SelfStart')
def self_start():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('SelfStart.html')

@app.route('/SignUp')
def sign_up():
    return render_template('SignUp.html')

@app.route('/SpecializedAutomation')
def specialized_automation():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('SpecializedAutomation.html')

@app.route('/SubjectMatterExpert')
def subject_matter_expert():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('SubjectMatterExpert.html')

@app.route('/Technology')
def technology():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Technology.html')

@app.route('/Telematics')
def telematics():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Telematics.html')

@app.route('/Templates')
def templates_page():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('Templates.html')

@app.route('/TermsToKnow')
def terms_to_know():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('TermsToKnow.html')

@app.route('/ToolKit')
def toolkit():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('ToolKit.html')

@app.route('/TransformationPlan')
def transformation_plan():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('TransformationPlan.html')

@app.route('/v2x')
def v2x():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('v2x.html')

@app.route("/getBotResponse")
def get_bot_response():
    userText = request.args.get('msg')
    queryAdderQry = request.args.get('addQry')
    queryAdderResp = request.args.get('addResp')
    if userText not in ["",None]:
        userText = " ".join(userText.strip().split())
        userText = html.escape(userText)
        userQueryLogTime = getCurrentTime()
        logDate = getCurrentDate()
        response = responseProcessor(userText).strip()
        if(response.startswith(unknownResponse)):
            logFolder = 'unknown'
        else:
            logFolder = 'known'
        responseLogTime = getCurrentTime()
        Responselogger(logFolder, "USER", userText, logDate, userQueryLogTime)
        Responselogger(logFolder, "BOT", response, logDate, responseLogTime)
    elif queryAdderResp not in ["",None] and queryAdderQry not in ["",None]:
        queryAdderQry = " ".join(queryAdderQry.strip().split())
        queryAdderResp = queryAdderResp.strip()
        queryAdderQry = html.escape(queryAdderQry)
        response = queryAdder(queryAdderQry,queryAdderResp)
    else:
        response = "Something Went Wrong, Please Contact the Admin"
    return html.escape(response)
    
@app.before_request
def before_request():
    if 'user' not in session and request.endpoint not in ['index', 'login']: 
        return redirect(url_for('index'))
    session.modified = True        

if __name__ == "__main__":
    app.run(threaded=True,port=5001)
