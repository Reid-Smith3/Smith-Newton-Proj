import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import tkinter as tk
from functools import partial
from itertools import permutations
import random
from PIL import ImageTk, Image

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# setting the id and ranges for the movie database
MOVIES_SPREADSHEET_ID = '1XrT1fRhhiS_Ga7yX9AEikkK5uFoGOfIIX6UpOXKLPLI'
name_range = 'B2:B158548'
year_range = 'C2:C158548'
genre_range = 'E2:E158548'
rating_range = 'H2:H158548'

# ranges for additional access
director_range = 'F2:F158548'
writer_range = 'G2:G158548'

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
fav_movies = ''
fav_movies_var = ''
user_list = []

# initializing the global variables for the data
g_values = ''
y_values = ''
n_values = ''
r_values = ''
d_values = ''
w_values = ''

# set creds to be a global variable to allow for access to
# other data/spreadsheets once the initial call has been made
creds = ''

def openSheet():
    """ Opens the movie database and returns the relevant information"""
    
    global creds
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
    result_ratings = sheet.values().get(spreadsheetId=MOVIES_SPREADSHEET_ID,
                                range=rating_range).execute()
    
    # modify global variables to avoid parameter passing
    # and ensure modification for future searches
    global g_values
    global n_values
    global y_values
    global r_values
    
    # get the actual values for each call
    g_values = result_genre.get('values', [])
    n_values = result_name.get('values', [])
    y_values = result_year.get('values', [])
    r_values = result_ratings.get('values', [])
    
    # load from local file
#     movie_file = open('moviedata.pickle', 'rb')
#     movie_dict = pickle.load(movie_file)
#     movie_file.close()

def openSheet2():
    """ Load the data for directors and writers to calculate
    correlation favorite movies from the user"""
    
    # Call the Sheets API
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    
    result_director = sheet.values().get(spreadsheetId=MOVIES_SPREADSHEET_ID,
                                range=director_range).execute()
    result_writer = sheet.values().get(spreadsheetId=MOVIES_SPREADSHEET_ID,
                                range=writer_range).execute()
    
    global d_values
    global r_values
    
    d_values = result_director.get('values', [])
    w_values = result_writer.get('values', [])
    
#     print('Sheet 2 Loaded')
    
def addGenre(button):
    """Add genres from each clicked button to the genre list"""
    genres.append(button)
    
def addMovie(button):
    """Add movies from each clicked button to the user list"""
    if button not in user_list:
        user_list.append(button)
    
def removeMovie(button):
    """Add movies from each clicked button to the user list"""
    bad = user_list.index(button)
    user_list.pop(bad)

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
        return upper
    except:
        return
    
def favoriteM(*args):
    """Detect a change in the favorite movies variable"""
    global fav_movies
    global fav_movies_var
    fav_movies = fav_movies_var.get()
    return fav_movies

def saveList(passed_list=user_list):
    """ Saves the user generated list to a local file"""
    save_file = open('savedlist.txt', 'w')
    for movie in passed_list:
        save_file.write('%s\n' % movie)
    
def info():
    """ Creates a separate UI detailing information about the
    background of the creators and their preferenes for movies,
    allowing the user to generate lsits from developer picks"""
    
    window_info = tk.Tk()
    
    # Reid's frame
    reid_frame = tk.Frame(master=window_info, bg='white')
    reid_label = tk.Label(master=reid_frame, text='About Reid', bg='white',
                    font=('Garamond',30,'bold'), fg='steel blue')
    reid_info = tk.Label(master=reid_frame, text='Use this space to give a brief background about ourselves, \n our contributions to the project, as well as our favorite movies.',
                    bg='white', font=('Garamond',14), fg='steel blue')
    reid_label.pack(side='top', expand=True)
    reid_info.pack(side='top', expand=True)
    reid_frame.pack(side='left', expand=True)
    
    # Christian's frame
    newt_frame = tk.Frame(master=window_info, bg='light grey')
    newt_label = tk.Label(master=newt_frame, text='About Christian', bg='light grey',
                    font=('Garamond',30,'bold'), fg='IndianRed4')
    newt_info = tk.Label(master=newt_frame, text='Use this space to give a brief background about ourselves, \n our contributions to the project, as well as our favorite movies.',
                    bg='light grey',font=('Garamond',14), fg='IndianRed4')
    newt_label.pack(side='top', expand=True)
    newt_info.pack(side='top', expand=True)
    newt_frame.pack(side='left', expand=True)
    
    # EASTER EGG CONTENT to generate lists of our
    # favorite movies here
    
    window_info.mainloop()

def advancedPref():
    """ Create the additional user interface for finetuning
    the user's preferences"""
    
    preferences = tk.Tk()
    preferences.configure(bg='white')
    
    # enter a favorite movie
    frame_movie = tk.Frame(master=preferences, bg='white')
    head_movie = tk.Label(master=frame_movie, text='Fan Favorites:', bg='white',
                    font=('Garamond',30,'bold'), fg='blue')
    info_movie = tk.Label(master=frame_movie, text='Enter 1 or more favorite movies, separating \n titles with commas.',
                    bg='white', font=('Garamond',14), fg='black')
    head_movie.pack(side='top')
    info_movie.pack(side='top')
    
    frame_movie.pack(side='top')
    
    # submission area for movie(s)
    frame_movie2 = tk.Frame(master=preferences, bg='white')
    label_movie = tk.Label(master=frame_movie2, text='Movies:', bg='white',
                    font=('Garamond',20,'bold'), fg='black')
    label_movie.pack(side='left')
    
    # set the variable for favorite movies
    global fav_movies_var
    fav_movies_var = tk.StringVar(frame_movie2)
    entry_movie = tk.Entry(master=frame_movie2, font=('Garamond',14),
                           textvariable=fav_movies_var, width=17)
    entry_movie.pack(side='left')
    fav_movies_var.trace("w", favoriteM) 
    
    frame_movie2.pack(side='top')
    
    # destroy the advanced preferences
    advanced_frame = tk.Frame(master=preferences, bg='white')
    advanced = tk.Button(master=advanced_frame, text='Submit Advanced',
                         command=preferences.destroy, relief='raised',
                       fg='red', font=('Garamond',14,'bold'))
    advanced.pack(side='bottom')
    advanced_frame.pack(side='bottom')
    preferences.mainloop()
    
def userList():
    """ Create the watch later list for the user"""
    
    # set up display
    window_display = tk.Tk()
    window_display.configure(bg='white')
    top_frame = tk.Frame(master=window_display, bg='white')
    text_frame = tk.Frame(master=top_frame, bg='white')
    
    # user did submit preferences, FIGURE OUT GOOD WAY TO DISPLAY
    if user_list:
        
        # iterate through items in list, create label
        # and button for each
        for i in range(len(user_list)):
            name_val = user_list[i][0]
            year_val = user_list[i][1]
            genres_val = user_list[i][2]
            corr = user_list[i][4]
            corr_val = round(corr, 6)
            label_num = i + 1
            
            # add text label for movie
            label_frame = tk.Frame(master=text_frame, bg='white', width=70)
            text_label = tk.Label(master=label_frame, text=str(label_num) + '. '
                            + str(name_val) + ': ' + str(year_val) + ': ' +
                            str(genres_val) + ': ' + str(corr_val) + '\n',
                            font=('Garamond', 14), bg='white',
                            anchor='w', width=66)
            text_label.pack(side='left')
            
            # add button for each movie
            button_movie = tk.Button(master=label_frame, text='-',
                       font=('Garamond',16,'bold'), fg='red3', height=1, width=4,
                    command=partial(removeMovie, user_list[i]), relief='raised')
            button_movie.pack(side='left')
            label_frame.pack(side='top')
    
    # force user to select genres to generate a movie list
    else:
        text_label = tk.Label(master=text_frame, text='No movies in watch later.',
                            bg='white', font=('Garamond',14))
        text_label.pack()
    text_frame.pack(side='top')
    top_frame.pack(side='top')
        
    # search again button
    bottom_frame = tk.Frame(master=window_display, bg='gray')
    search_button = tk.Button(master=bottom_frame, text='Refresh List',
                        command=lambda:[window_display.destroy(), userList()],
                       relief='raised', fg='blue', font=('Garamond',16,'bold'))
    search_button.pack(side='right', expand=True)
    
    # close current button
    exit_button = tk.Button(master=bottom_frame, text='Exit List',
                        command=window_display.destroy,
                       relief='raised', fg='red', font=('Garamond',16,'bold'))
    exit_button.pack(side='left', expand=True) 
    
    # save list button
    save_button = tk.Button(master=bottom_frame, text='Save List',
                        command=saveList, relief='raised',
                        fg='dark green', font=('Garamond',16,'bold'))
    save_button.pack(side='left', expand=True)
    bottom_frame.pack(side='bottom')
    
    window_display.mainloop()

def initializeUI():
    """ Initialize the user interface"""
    
    window = tk.Tk()
    window.configure(bg='white')
    
    # title of application
    frame_title = tk.Frame(master=window, bg='white')
    title = tk.Label(master=frame_title, text='The GOAT Movie Recommendation', bg='white',
                    font=('Garamond',36,'bold'), fg='steel blue')
    title.pack(side='top')
    title_info = tk.Label(master=frame_title, text='Brought to you by: Christian Newton and Reid Smith',
                    bg='white', font=('Garamond',16,'bold'), fg='black')
    title_info.pack(side='top')
    title_info2 = tk.Label(master=frame_title, text='Sponsored by: Philip Caplan, Esq.',
                    bg='white', font=('Garamond',12), fg='gray')
    title_info2.pack(side='top')
    frame_title.pack(side='top')
    
    # load image (CHANGE TO LOGO)
    img = ImageTk.PhotoImage(Image.open('newtreid.jpg'))
    panel = tk.Label(master=window, image=img, bg='white')
    panel.pack(side='top', fill='both',expand=True)
    
    # button performs 3 tasks: destroys the window, loads info
    # from database, and opens main UI
    bottom_frame = tk.Frame(master=window, bg='white')
    button_search = tk.Button(master=bottom_frame, text='Search!',
                         command=lambda:[window.destroy(), openSheet(), mainUI()],
                      relief='raised', fg='red', font=('Garamond',14,'bold'))
    button_search.pack(side='right', expand=True)
    
    # info button for creators and easter egg movie lists
    button_info = tk.Button(master=bottom_frame, text='Info',
                    command=info, relief='raised', fg='black',
                    font=('Garamond',14,'bold'))   
    button_info.pack(side='left', expand=True)
    
    bottom_frame.pack(side='bottom', expand=True)
    
    window.mainloop()

def mainUI():
    """ Create the layout for the main user interface"""

    window = tk.Tk()
    
    # genres 
    frame_genre = tk.Frame(master=window, width=600, height=800,
                      bg='white')
    head_genre = tk.Label(master=frame_genre, text='Genres:', bg='white',
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
                    command=partial(addGenre, db_genres[i]), relief='raised', bg='white')
        button_genre.pack(side='top')
    # relief: flat, groove, raised, ridge, solid, sunken
    
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

    year_menu1 = tk.OptionMenu(start_frame, year_start_var, *option_list,)
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
    head_rating.pack(side='top')
    info_rating1.pack(side='top')
    info_rating2.pack(side='top')
    
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
    
    # upper rating entry
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
    bottom_frame = tk.Frame(master=window, bg='white')
    
    enter = tk.Button(master=bottom_frame, text='Submit Preferences',
                         command=window.destroy, relief='raised',
                       fg='red', font=('Garamond',14,'bold'))
    enter.pack(side='right')
    
    # allow user to acces info button in mainUI also
    info_button = tk.Button(master=bottom_frame, text='Info',
                         command=info, relief='raised',
                       fg='black', font=('Garamond',14,'bold'))
    info_button.pack(side='left')
    
    # view list button
    view_list = tk.Button(master=bottom_frame, text='View List',
                         command=userList, relief='raised',
                       fg='dark green', font=('Garamond',14,'bold'))
    view_list.pack(side='left')
    
    # advanced preferences button, load next set of data
    advanced = tk.Button(master=bottom_frame, text='Advanced',
                         command=lambda:[openSheet2(), advancedPref()], relief='raised',
                       fg='blue', font=('Garamond',14,'bold'))
    advanced.pack(side='left')
    
    bottom_frame.pack(side='bottom')
    
    # how to set the background for the whole window
    window.configure(bg='white')
    window.mainloop()
    
def createList():
    """ Performs the search of the database by matching
    the user's input to the available movies, also
    restarts the UI if the user desires to search again"""
    
    # clean the favorite movies input from the advanced
    # preferences section (USE THIS SOMEHOW)
    if fav_movies:
        split_movies = fav_movies.split(',')
        final_movies = []
        for mov in split_movies:
            strip_mov = mov.strip()
            mov_words = strip_mov.split(' ')
            cap = [w.capitalize() for w in mov_words]
            title = ' '.join(cap)
            final_movies.append(title)
    
#     print(genres)
#     print(year_start, year_end)
#     print(lower, upper)
#     print(final_movies)

    # first select movies that match the year and rating
    # range and package the relevant data into a single list
    year_subset = []
    count = 0
    for movie_year in y_values:
        if movie_year[0] != '\\N':
            yr = int(movie_year[0])
            rating = float(r_values[count][0])
            if (yr >= year_start and yr <= year_end) and (rating >= lower and rating <= upper):
                local_genres = g_values[count]
                name = n_values[count][0]
                year_subset.append([name, yr, local_genres, rating])
        count += 1
            
    # load pre-processed correlation dictionary
    corr_file = open('correlation.pickle', 'rb')
    corr_dict = pickle.load(corr_file)
    corr_file.close()

    # calculate genre correlations for each movie and give
    # each movie a unique key to later perform sort operations
    corr_subset = {}
    count = 0
    
    # use the maximum correlation value for matching genres
    max_corrval = 1.5 * max(corr_dict.values())
    for movie in year_subset:
        movie_genres = movie[2][0]
        search_genres = movie_genres.split(',')
        corr_val = 0
        local_genres = genres.copy()
        
        # decrease correlation when multiple genres match
        alpha = 1
        
        # decrease correlation for triple genre keys
        beta = len(search_genres)
        
        # decrease correlation for more triple genres comparisons
        gamma = len(local_genres)
        
        # paired genre correlation
        for movie_gen in search_genres:   
            for gen in local_genres:
                test_1 = (movie_gen, gen)
                test_2 = (gen, movie_gen)
                if movie_gen == gen:
                    
                    # decreases when multiple genres match
                    corr_val += max_corrval / alpha
                    alpha += 1
                elif test_1 in corr_dict:
                    corr_val += corr_dict[test_1]
                elif test_2 in corr_dict:
                    corr_val += corr_dict[test_2]
        
        # triple genre correlation with two from database
        if beta == 2:
            g1 = search_genres[0]
            g2 = search_genres[1]
            
            # scale permutations with 2 genres
            delta = beta * gamma
            
            # permute all possible keys
            for gen in local_genres:
                perm_twogen = list(permutations([g1,g2,gen]))
                for perm in perm_twogen:
                    if perm in corr_dict:
                        corr_val += corr_dict[perm] / delta
        
        # triple genre correlation with three from database
        if beta == 3:
            g1 = search_genres[0]
            g2 = search_genres[1]
            g3 = search_genres[2]
            
            # scale permutations with 3 genres
            delta = beta * beta * gamma
            
            # permute all possible keys
            for gen in local_genres:
                perm_threegen_1 = list(permutations([g1,g2,gen]))
                perm_threegen_2 = list(permutations([g1,g3,gen]))
                perm_threegen_3 = list(permutations([g2,g3,gen]))
                for perm in perm_threegen_1:
                    if perm in corr_dict:
                        corr_val += corr_dict[perm] / delta
                for perm in perm_threegen_2:
                    if perm in corr_dict:
                        corr_val += corr_dict[perm] / delta
                for perm in perm_threegen_3:
                    if perm in corr_dict:
                        corr_val += corr_dict[perm] / delta
                             
        # scale final correlation value by alpha and beta
        corr_val = alpha * corr_val / beta
        corr_val = round(corr_val, 6)
        movie.append(corr_val)
        corr_subset[count] = movie
        count += 1
    
    # sort dictionary based on correlation value
    corr_sorted = sorted(corr_subset.items(), key=lambda item: item[1][4], reverse=True)
    
    # sort dictionary based on rating
    rating_sorted = sorted(corr_subset.items(), key=lambda item: item[1][3], reverse=True)
    
    # take one random element from rating and add to final
    real_values = []
    total = 1
    mid_rank = random.choice(rating_sorted)
    real_values.append(mid_rank[1])
    
    # get an even spread of years
    year_range = year_end - year_start
    half_decades = year_range // 5
    if half_decades == 0:
        half_decades = 1

    # find the top values from each five year period
    # within the user selected range of years
    halfd_sorted = year_sorted.copy()
    temp_sorted = year_sorted.copy()
    year_corr = []
    
    # establish a cutoff for the top values
    cut_val = max_corrsort / 10
    year_begin = halfd_sorted[0][1][1]
    for i in range(half_decades):
        half_list = []
        
        # iterate through all movies currently in list
        for sort in halfd_sorted:
            if sort[1][1] - 5 < year_begin:
                if sort[1][4] + cut_val >= max_corrsort:
                    half_list.append(sort)
                    
            # move onto next period
            else:
                year_begin += 5
                break
            temp_sorted.pop(0)
        halfd_sorted = temp_sorted.copy()
        year_corr.append(half_list)
    
    # if less than 20, 15, or 5 available, set max
    max_corr = len(corr_sorted)
    max_rating = len(rating_sorted)
    max_total = max_corr + max_rating
    if max_total >= 15:
        max_total = 15
    if max_corr >= 10:
        max_corr = 10
    if max_rating >= 4:
        max_rating = 4
        
    # how many movies to take from each year range
    # when there are fewer year ranges than movies taken
    corr_cycle = max_corr // half_decades
    index_count = 0
    
    # guarantee 20 items are taken, intersection of lists
    corr_count = 0
    rating_count = 0
                
    while total < max_total:
        
        # ensure a balance of correlation to ratings
        if corr_count < max_corr:
            if half_decades <= max_corr:
                
                for i in range(corr_cycle):
                    corr_extract = random.choice(year_corr[index_count])
                    corr_id = corr_extract[0]
                    corr_match = False
                    for title in real_values:
                        if title[0] == corr_id:
                            corr_match = True
                            ind = year_corr[index_count].index(corr_extract)
                            year_corr[index_count].pop(ind)
                    
                    if not corr_match:
                        real_values.append(corr_extract[1])
                        ind = year_corr[index_count].index(corr_extract)
                        year_corr[index_count].pop(ind)
                        corr_count += 1
                        total += 1
                        
                        if corr_count == max_corr or total == max_total:
                            break
                    
                index_count += 1
                if index_count == half_decades:
                    index_count = 0
                    
            else:
                corr_cyclelist = [ind for ind in range(len(year_corr))]
                random_index = random.choice(corr_cyclelist)
                corr_extract = random.choice(year_corr[random_index])
                corr_id = corr_extract[0]
                corr_match = False
                for title in real_values:
                    if title[0] == corr_id:
                        corr_match = True
                        year_corr.pop(random_index)
                        
                if not corr_match:
                    real_values.append(corr_extract[1])
                    year_corr.pop(random_index)
                    corr_count += 1
                    total += 1
            
        if total == max_total:
            break
        
        # ensure a balance of rating to correlation
        if rating_count < max_rating:
            rating_id = rating_sorted[0][0]
            rating_match = False
            for title in real_values:
                if title[0] == rating_id:
                    rating_match = True
                    rating_sorted.pop(0)
                    
            if not rating_match:
                real_values.append(rating_sorted[0][1])
                rating_sorted.pop(0)
                rating_count += 1
                total += 1
                
    # shuffle final list to create nuanced feel for user
    random.shuffle(real_values)
    
    # set up display
    window_display = tk.Tk()
    window_display.configure(bg='white')
    top_frame = tk.Frame(master=window_display, bg='white')
    text_frame = tk.Frame(master=top_frame, bg='white')
    
    # user did submit preferences, FIGURE OUT GOOD WAY TO DISPLAY
    if real_values and genres:

        # iterate through items in list, create label
        # and button for each
        for i in range(len(real_values)):
            name_val = real_values[i][0]
            year_val = real_values[i][1]
            genres_val = real_values[i][2]
            corr = real_values[i][4]
            corr_val = round(corr, 6)
            label_num = i + 1
            
            # add text label for movie
            label_frame = tk.Frame(master=text_frame, bg='white', width=70)
            text_label = tk.Label(master=label_frame, text=str(label_num) + '. '
                            + str(name_val) + ': ' + str(year_val) + ': ' +
                            str(genres_val) + ': ' + str(corr_val) + '\n',
                            font=('Garamond', 14), bg='white',
                            anchor='w', width=66)
            text_label.pack(side='left')
            
            # add button for each movie
            button_movie = tk.Button(master=label_frame, text='+',
                       font=('Garamond',16,'bold'), fg='dark green', height=1, width=4,
                    command=partial(addMovie, real_values[i]), relief='raised')
            button_movie.pack(side='top')
            label_frame.pack(side='top')

    
    # force user to select genres to generate a movie list
    else:
        text_label = tk.Label(master=text_frame, text='No movies found.',
                            bg='white', font=('Garamond',14))
        text_label.pack()
    text_frame.pack(side='top')
    top_frame.pack(side='top')
        
    # search again button
    bottom_frame = tk.Frame(master=window_display, bg='gray')
    search_button = tk.Button(master=bottom_frame, text='Search Again',
                        command=lambda:[window_display.destroy(), subsequent()],
                       relief='raised', fg='blue', font=('Garamond',16,'bold'))
    search_button.pack(side='right', expand=True)
    
    # close current button
    exit_button = tk.Button(master=bottom_frame, text='Exit Search',
                        command=window_display.destroy,
                       relief='raised', fg='red', font=('Garamond',16,'bold'))
    exit_button.pack(side='left', expand=True)
        
    # save list button
    save_button = tk.Button(master=bottom_frame, text='Save List',
                        command=lambda:[saveList(real_values)],
                       relief='raised', fg='dark green', font=('Garamond',16,'bold'))
    save_button.pack(side='left', expand=True)
    bottom_frame.pack(side='bottom')
    
    window_display.mainloop()

def first():
    """Starts the process of manufacturing a list from a user's
    inputted preferences."""
    
    # initialize the UI for the first search
    initializeUI()
    
    # obtain a curated list of movies and displays, also gives
    # the option of reinitializing the UI if the user wants
    # to search again
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
    
    global fav_movies
    fav_movies = ''
  
    # do not go back to the main screen for subsequent runs
    mainUI()
    
    createList()
    
if __name__ == '__main__':
    first()
