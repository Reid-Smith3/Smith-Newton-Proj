import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import tkinter as tk
from functools import partial

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# setting the id and ranges for the movie database
MOVIES_SPREADSHEET_ID = '1XrT1fRhhiS_Ga7yX9AEikkK5uFoGOfIIX6UpOXKLPLI'
name_range = 'B2:B158548'
genre_range = 'E2:E158548'
year_range = 'C2:C158548'
rating_range = 'H2:H158548'

# initializing the global variables for the UI
genres = []
year_start = 1930
year_end = 2020
year_start_var = ''
year_end_var = ''
lower = 0
lower_var = ''
upper = 10
upper_var = ''

# initializing the global variables for the data
g_values = ''
y_values = ''
n_values = ''
r_values = ''

def openSheet():
    """ Opens the movie database and returns the relevant information"""
    
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when authorization flow completes for the first time
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_genre = sheet.values().get(spreadsheetId=MOVIES_SPREADSHEET_ID,
                               range=genre_range).execute()
    result_name = sheet.values().get(spreadsheetId=MOVIES_SPREADSHEET_ID,
                                range=name_range).execute()
    result_year = sheet.values().get(spreadsheetId=MOVIES_SPREADSHEET_ID,
                                range=year_range).execute()
    
    # modify global variables to avoid parameter passing
    # and ensure modification for future searches
    global g_values
    global n_values
    global y_values
    
    # get the actual values for each call
    g_values = result_genre.get('values', [])
    n_values = result_name.get('values', [])
    y_values = result_year.get('values', [])

#     movie_file = open('moviedata.pickle', 'rb')
#     movie_dict = pickle.load(movie_file)
#     movie_file.close()

    return g_values, n_values, y_values

def addGenre(button):
    """Add genres from each clicked button to the genre list"""
    genres.append(button)

def callbackS(*args):
    """Detect a change in the year start variable"""
    global year_start
    global year_start_var
    year_start = year_start_var.get()
    return year_start

def callbackE(*args):
    """Detect a change in the year end variable"""
    global year_end
    global year_end_var
    year_end = year_end_var.get()
    return year_end

def ratingL(*args):
    """Detect a change in the lower rating variable"""
    global lower
    global lower_var
    
    # to prevent an error if the user types a value and
    # subsequently changes it
    try:
        lower = lower_var.get()
        print(lower)
        return lower
    except:
        return

def ratingU(*args):
    """Detect a change in the upper rating variable"""
    global upper
    global upper_var
    
    # to prevent an error if the user types a value and
    # subsequently changes it
    try:
        upper = upper_var.get()
        print(upper)
        return upper
    except:
        return

def initializeUI():
    """ Initialize the user interface"""

    window = tk.Tk()
    
    # genres 
    frame_genre = tk.Frame(master=window, width=600, height=800,
                      bg='white')
    head_genre = tk.Label(master=frame_genre, text='Genre(s):', bg='white',
                    font=('Garamond',30,'bold'), fg='blue')
    info_genre = tk.Label(master=frame_genre, text='Select 1 or more of the listed genres.',
                    bg='white', font=('Garamond',14), fg='black')
    head_genre.pack(side='top')
    info_genre.pack(side='top')
    
    # create all available genre buttons for user to select from
    db_genres = ['Horror','Comedy','Romance','Sci-Fi', 'Drama',
                 'Action','Adventure','Crime','War','Biography',
                 'Musical','Fantasy','Western','Thriller','Documentary',
                 'Film-Noir']
    genre_colors = ['IndianRed4','steel blue','purple4','slate gray','blue4',
                    'red3','dark green','DarkOrange2']
    
    for i in range(len(db_genres)):
        if i % 4 == 0:
            button_frame = tk.Frame(master=frame_genre, bg='white')
        elif i % 4 == 3:
            button_frame.pack(side='left')
        button_genre = tk.Button(master=button_frame, text=db_genres[i],
                       font=('Garamond',16), fg=genre_colors[i%8],
                    command=partial(addGenre, db_genres[i]), relief='groove', bg='white')
        button_genre.pack(side='top')
    
    frame_genre.pack(side='top',expand=True)
    
    # year
    frame_year = tk.Frame(master=window, width=400, height=300,
                      bg='white')
    head_year = tk.Label(master=frame_year, text='Year Selection:', bg='white',
                    font=('Garamond',30,'bold'), fg='blue')
    info_year = tk.Label(master=frame_year,
                    text='Enter a lower and/or upper bound for the desired release \n years of movies.',
                    bg='white', font=('Garamond',14), fg='black')
    head_year.pack(side='top')
    info_year.pack(side='top')
    
    # start year section
    start_frame = tk.Frame(master=frame_year, bg='white')
    
    # global variable to modify the start year variable from dropdown
    global year_start_var
    year_start_var = tk.IntVar(frame_year)
    year_start_var.set(1930)
    
    start_year = tk.Label(master=start_frame, text='Start Year:', bg='white',
                    font=('Garamond',20,'bold'), fg='black')
    start_year.pack(side='left', expand=True)
    
    # create list of possible years accessible in database
    option_list = []
    start = 1930
    for i in range(91):
        option_list.append(str(i+start))

    year_menu1 = tk.OptionMenu(start_frame, year_start_var, *option_list)
    year_menu1.pack(side='left')    
    year_start_var.trace("w", callbackS)    
    start_frame.pack(side='left')
    
    # end year section
    end_frame = tk.Frame(master=frame_year, bg='white')
    
    # global variable to modify end year variable from dropdown
    global year_end_var
    year_end_var = tk.IntVar(frame_year)
    year_end_var.set(2020)
    
    end_year = tk.Label(master=end_frame, text='End Year:', bg='white',
                    font=('Garamond',20,'bold'), fg='black')
    end_year.pack(side='left', expand=True)

    year_menu2 = tk.OptionMenu(end_frame, year_end_var, *option_list)
    year_menu2.pack(side='left')
    year_end_var.trace("w", callbackE)
    end_frame.pack(side='right')
    
    frame_year.pack(side='top')
    
    # rating
    frame_rating = tk.Frame(master=window, width=400, height=300,
                      bg='white')
    head_rating = tk.Label(master=frame_rating, text='Rating Range:', bg='white',
                    font=('Garamond',30,'bold'), fg='blue')
    info_rating1 = tk.Label(master=frame_rating,
                text='Optional Preference',
                bg='white', font=('Garamond',16,'bold'), fg='black')
    info_rating2 = tk.Label(master=frame_rating,
                text='Enter numbers between 1 and 10 for the lower and upper \n bounds of crowd-sourced ratings.',
                bg='white', font=('Garamond',14), fg='black')
    head_rating.pack(side='top', expand=False)
    info_rating1.pack(side='top', expand=False)
    info_rating2.pack(side='top', expand=False)
    
    # lower rating entry
    lower_frame = tk.Frame(master=frame_rating, bg='white')
    lower_label = tk.Label(master=lower_frame, text='Lower:', bg='white',
                    font=('Garamond',20,'bold'), fg='black')
    lower_label.pack(side='left', expand=True)
    
    # set variable for lower rating
    global lower_var
    lower_var = tk.DoubleVar(lower_frame)
    lower_var.set('')
    lower_entry = tk.Entry(master=lower_frame, font=('Garamond',20),
                           textvariable=lower_var, width=7)
    lower_var.trace('w', ratingL)
    lower_entry.pack(side='left')
    lower_frame.pack(side='left')
    
    # upperlower rating entry
    upper_frame = tk.Frame(master=frame_rating, bg='white')
    upper_label = tk.Label(master=upper_frame, text='Upper:', bg='white',
                    font=('Garamond',20,'bold'), fg='black')
    upper_label.pack(side='left', expand=True)
    
    # set variable for upper entry
    global upper_var
    upper_var = tk.DoubleVar(upper_frame)
    upper_var.set('')
    upper_entry = tk.Entry(master=upper_frame, font=('Garamond',20),
                           textvariable=upper_var, width=7)
    upper_var.trace('w', ratingU)
    upper_entry.pack(side='left')
    upper_frame.pack(side='right')
    
    frame_rating.pack(side='top', expand=True)
    
    # destroy button
    button = tk.Button(master=window, text='Enter Preferences',
                         command=window.destroy, relief='raised',
                       fg='red', font=('Garamond',14,'bold'))
    button.pack(side='bottom')
    window.mainloop()
    
def createList():
    """ Performs the search of the database by matching
    the user's input to the available movies, also
    restarts the UI if the user desires to search again"""
    print(lower, upper)

    # add movies to different genre lists
    final_values = []
    i = 0
    for gen in genres:
        final_genre = []
        for movie in g_values:
            movie_genres = movie[0].split(',')
            for movie_gen in movie_genres:
                if gen == movie_gen and y_values[i][0] != '\\N':                    
                    final_genre.append([n_values[i], y_values[i]])
            i += 1
        final_values.append(final_genre)
        i = 0
    
    # select movies that match year range
    real_values = []
    for sublist in final_values:
        real_list = []
        for entry in sublist:
            yr = int(entry[1][0])
            if yr >= year_start and yr <= year_end:
                real_list.append(entry)
        real_values.append(real_list)

    # display separate lists for each genre (CHANGE)
    window_display = tk.Tk()
    for i in range(len(real_values)):
        text_box = tk.Text(master=window_display, wrap='none',
                           font=('Garamond', 12))
        entry_num = len(real_values[i]) - 1
        for j in range(len(real_values[i])):
            row = entry_num - j
            name_val = real_values[i][row][0][0]
            year_val = real_values[i][row][1][0]
            label_num = row + 1
            text_box.insert(1.0, str(label_num) + '. ' +
                            str(name_val) + ': ' + str(year_val) + '\n')
        text_box.pack(side='top')
        text_box.insert(1.0, str(genres[i]) + ': Year \n\n')
    
    # user did not submit any preferences, avoid error
    if not real_values:
        text_box = tk.Text(master=window_display, wrap='none',
                           font=('Garamond', 12))
        text_box.insert(1.0, 'No data found.')
        text_box.pack(side='top')
    
    # search again button
    button = tk.Button(text='Search Again',
                        command=subsequent,
                       relief='raised', fg='blue', font=('Garamond',16,'bold'))
    button.pack(side='right', expand=True)
    
    # close current button
    button = tk.Button(text='Exit Search',
                        command=window_display.destroy,
                       relief='raised', fg='red', font=('Garamond',16,'bold'))
    button.pack(side='left', expand=True)
    
    window_display.mainloop()

def first():
    """Runs the process of manufacturing a list from a user's
    inputted preferences."""
    
    # open the sheet and obtain values
    openSheet()
    
    # initialize the UI for the first search
    initializeUI()
    
    # obtain a curated list of movies and displays, also
    # reinitializes the UI if the user wants to search again
    createList()
    
def subsequent():
    """ If the user searches again, only initialize UI
    and create list rather than loading spreadsheet"""
    
    # need to reset the global variables
    global genres
    genres = []
    
    global year_start
    year_start = 1930
    global year_end
    year_end = 2020
    
    global lower
    lower = 0
    global upper
    upper = 10
  
    initializeUI()
    
    createList()
    
if __name__ == '__main__':
    first()
