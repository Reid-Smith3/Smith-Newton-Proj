import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
MOVIES_SPREADSHEET_ID = '1N_mLLiXweGWVjdVoHCOo8lCbHMP96K0bA7M62M5kK54'
genre_range = 'F2:F158548'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    
#     creds = None
#     # The file token.pickle stores the user's access and refresh tokens, and is
#     # created automatically when the authorization flow completes for the first
#     # time.
#     if os.path.exists('token.pickle'):
#         with open('token.pickle', 'rb') as token:
#             creds = pickle.load(token)
#     # If there are no (valid) credentials available, let the user log in.
#     if not creds or not creds.valid:
#         if creds and creds.expired and creds.refresh_token:
#             creds.refresh(Request())
#         else:
#             flow = InstalledAppFlow.from_client_secrets_file(
#                 'credentials.json', SCOPES)
#             creds = flow.run_local_server(port=0)
#         # Save the credentials for the next run
#         with open('token.pickle', 'wb') as token:
#             pickle.dump(creds, token)
# 
#     service = build('sheets', 'v4', credentials=creds)
# 
#     # Call the Sheets API
#     sheet = service.spreadsheets()
#     result_genre = sheet.values().get(spreadsheetId=MOVIES_SPREADSHEET_ID,
#                                 range=genre_range).execute()
#     
#     # get the actual values for each call
#     g_values = result_genre.get('values', [])

    g_file = open('gvals.pickle', 'rb')
    g_values = pickle.load(g_file)
    g_file.close()
    
    # tally paired genres
    genre_dict = {}
    for movie in g_values:
        movies = movie[0].split(',')
  
        # calculate genre pairings
        for gen_1 in movies:
            for gen_2 in movies:
                if gen_1 != gen_2:
                    if ((gen_1, gen_2) or (gen_2, gen_1)) not in genre_dict:
                        genre_dict[(gen_1, gen_2)] = 1
                    else:
                        genre_dict[(gen_1, gen_2)] += 1
                        
            # remove the genre from the list for each movie
            ind = movies.index(gen_1)
            movies.pop(ind)
        
        # if movie has 3 genres, calculate triple correlation
        movies_2 = movie[0].split(',')
        if len(movies_2) == 3:
            key = (movies_2[0], movies_2[1], movies_2[2])
            if key not in genre_dict:
                genre_dict[key] = 1
            else:
                genre_dict[key] += 1
    
    # fix flipped identical keys
    dict_1 = genre_dict.copy()
    dict_2 = genre_dict.copy()
    for pair_1 in dict_1:
        for pair_2 in dict_2:
            
            # only fix identical keys if a pair, not triple
            if len(pair_1) < 3 and len(pair_2) < 3:
                if (pair_1[0] == pair_2[1] and pair_1[1] == pair_2[0]) and (pair_1 in genre_dict and pair_2 in genre_dict):
                    genre_dict[pair_1] += genre_dict[pair_2]
                    genre_dict.pop(pair_2, None)

    # calculate ratio of correlations to total movies
    corr_dict = {}
    num_movies = 158548
    for pair in genre_dict:
        fraction = genre_dict[pair]/num_movies
        rounded = round(fraction, 6)
        corr_dict[pair] = rounded
        
        # increase value of triple keys
        if len(pair) == 3:
            corr_dict[pair] = corr_dict[pair]

    # save to file for use in main script
    corr_file = open('correlation.pickle', 'wb')
    pickle.dump(corr_dict, corr_file)
    corr_file.close()
        
if __name__ == '__main__':
    main()
