import json
import requests as requests
import os.path
import xlsxwriter
import urllib.request
import io

# firstly, fetch data using the PuppetPlays API
url = 'https://api.puppetplays.eu/graphql/'
query = """{
  entries(section: "works") {
    title
    authors {
      title
    }
    mostRelevantDate
    mainImage @transform(height: 100) {
      id
      url
      filename
      ... on images_Asset {
        title
        alt
        description
        copyright
      }
    }
    medias @transform(height: 100) {
      id
      url
      ... on images_Asset {
        title
        alt
        filename
        description
        copyright
      }
    }
  }
}

"""
r = requests.post(url, json={'query': query})

# load data and transform it
json_data = json.loads(r.text)
liste_oeuvres = json_data['data']['entries']

# creating directory for both Excel and images repertoire
directory = "liste_image_oeuvres"
if not os.path.exists(directory):
    os.makedirs(directory)
if not os.path.exists(directory + "\\image_oeuvre"):
    os.makedirs(directory + "\\image_oeuvre")

header = ['Titre de l\'oeuvre', 'Date', 'Auteurs', 'Image', 'Titre de l\'image', 'Nom du fichier', 'Description',
          'Copyright', 'Alt Text', 'URL de l\'image']  # 9 fields

# Init Excel Worksheet + change size of cells
workbook = xlsxwriter.Workbook('.\\' + directory + '\\image_oeuvres.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(3, 3, 100)
worksheet.set_default_row(250)
worksheet.set_row(0, 20)
text_format = workbook.add_format({'text_wrap': True})
worksheet.set_column(0, 2, 25, text_format)
worksheet.set_column(4, 8, 25, text_format)

# I create a table in Excel, i get the length of the result in order to create a suitable table.
taille_tableau = len(liste_oeuvres)
limit = 'A1:J' + str(taille_tableau)
worksheet.add_table(limit, {'columns': [{'header': 'Titre de l\'oeuvre'}, {'header': 'Date'},
                                        {'header': 'Auteurs'}, {'header': 'Image'},
                                        {'header': 'Titre de l\'image'}, {'header': 'Nom du fichier'},
                                        {'header': 'Description'}, {'header': 'Copyright'}, {'header': 'Alt Text'}, {'header': 'URL de l\'image'}]})

row = 0
col = 0
#worksheet.write_row(0, 0, header)  # write the first line of the table with headers, and add a newline to not overwrite it
row += 1
i = 0  # i and j are both variables sets to know how many works have main image
j = 0
oeuvres_sans_images = []

for work in liste_oeuvres:
    liste_auteurs = []

    try:
        mainImageUrl = work['mainImage'][0]['url']
        worksheet.write(row, col + 9, mainImageUrl)

        mainImage = io.BytesIO(urllib.request.urlopen(mainImageUrl).read())
        filename = work['mainImage'][0]['filename']
        worksheet.insert_image(row, col + 3, url,
                               {'image_data': mainImage, 'object_position': 1, 'x_scale': 0.5, 'y_scale': 0.5})
        img_data = requests.get(mainImageUrl).content
        if not os.path.exists(directory + "\\image_oeuvre\\" + filename):
            with open(directory + "\\image_oeuvre\\" + filename, 'wb') as handler:
                handler.write(img_data)

        mainImageTitle = work['mainImage'][0]['title']
        worksheet.write(row, col + 4, mainImageTitle)

        worksheet.write(row, col + 5, filename)

        mainImageDesc = work['mainImage'][0]['description']
        worksheet.write(row, col + 6, mainImageDesc)

        mainImageCopyright = work['mainImage'][0]['copyright']
        worksheet.write(row, col + 7, mainImageCopyright)

        mainImageAlt = work['mainImage'][0]['alt']
        worksheet.write(row, col + 8, mainImageAlt, text_format)
        i += 1
    except:
        j += 1
        oeuvres_sans_images.append(work['title'])
        if work['medias']:
            try:
                mainImageUrl = work['mainImage'][0]['url']
                print(work['title'] + " : image mal nommée ")
                print(mainImageUrl)
            except:
                print(work['title'] + " : image mal placée ")
                pass

        continue

    titleWork = work['title']
    worksheet.write(row, col, titleWork)

    date = work['mostRelevantDate']
    worksheet.write(row, col + 1, date)

    for auteurs in work['authors']:
        liste_auteurs.append(auteurs['title'])
    authors = ', '.join(liste_auteurs)
    worksheet.write(row, col + 2, authors)

    if work['medias']: # write a row for each media
        for media in work['medias']:
            row += 1
            worksheet.write(row, col, titleWork)
            worksheet.write(row, col + 1, date)
            worksheet.write(row, col + 2, authors)
            mediaImageUrl = media['url']
            worksheet.write(row, col + 9, mediaImageUrl)

            mediaImage = io.BytesIO(urllib.request.urlopen(mediaImageUrl).read())
            mediafilename = media['filename']
            worksheet.insert_image(row, col + 3, url,
                                   {'image_data': mediaImage, 'object_position': 1, 'x_scale': 1, 'y_scale': 1,
                                    'x_offset': 15})
            img_data = requests.get(mediaImageUrl).content
            if not os.path.exists(directory + "\\image_oeuvre\\" + mediafilename):
                with open(directory + "\\image_oeuvre\\" + mediafilename, 'wb') as handler:
                    handler.write(img_data)
            worksheet.write(row, col + 5, mediafilename)
            mediaImageTitle = media['title']
            worksheet.write(row, col + 4, mediaImageTitle)

            mediaImageDesc = media['description']
            worksheet.write(row, col + 6, mediaImageDesc)

            mediaImageCopyright = media['copyright']
            worksheet.write(row, col + 7, mediaImageCopyright)

            mediaImageAlt = media['alt']
            worksheet.write(row, col + 8, mediaImageAlt, text_format)
    row += 1



workbook.close()
z = i + j
print("Il y a " + str(i) + " oeuvres avec une image sur un total de " + str(z) + " oeuvres. Il y a donc " + str(
    j) + " oeuvres sans image.")
print("Voici la liste des oeuvres sans images : ", oeuvres_sans_images)
