# importing the requests library 
import requests 
import xlsxwriter 
# api-endpoint 
name=input("Enter Name Of Movie")
link = "https://api.themoviedb.org/3/search/movie?api_key=1dcf69b9b95240032c80e5d374ca2bee&language=en-US&query="+name

movie_name=[]
movie_overview=[]
movie_release=[]
# sending get request and saving the response as response object 
r = requests.get(url = link) 
data = r.json() 
results=data['total_results']
for i in range(0,5):
    movie_name.append(data['results'][i]['original_title'])
    movie_release.append(data['results'][i]['release_date'])
    movie_overview.append(data['results'][i]['overview'])

workbook = xlsxwriter.Workbook('movie.xlsx') 
worksheet = workbook.add_worksheet() 
worksheet.write('A1', 'Movie Name') 
worksheet.write('B1', 'Movie Release Date') 
worksheet.write('C1', 'Movie Description') 
for i in range(2,5):
    wb='A'+str(i)
    wbb='B'+str(i)
    wbbb='C'+str(i)
    worksheet.write(wb, movie_name[i-2]) 
    worksheet.write(wbb, movie_release[i-2]) 
    worksheet.write(wbbb, movie_overview[i-2]) 
# Finally, close the Excel file 
# via the close() method. 
workbook.close() 
