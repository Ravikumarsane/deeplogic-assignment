import requests    
 
def topnews():
    query_params = {
      "source": "time",
      "sortBy": "top",
      "apiKey": "4dbc17e007ab436fb66416009dfb59a8"
    }
    main_url = " https://newsapi.org/v1/articles"
 

    res = requests.get(main_url, params=query_params)
    open_time_page = res.json()
 

    article = open_time_page["articles"]
 

    results = []
    links= []
     
    for ar in article:
        results.append(ar["title"])
        links.append(ar["url"])
         
    for i in range(0,6):
         

        print(i + 1, results[i],"\n Link:",links[i],"\n")
 

    from win32com.client import Dispatch
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(results)                
 

if __name__ == '__main__':
     

    topnews()
