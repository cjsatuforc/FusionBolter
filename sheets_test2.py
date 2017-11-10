import json
from urllib import request, error
def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """



    query = ''
    spreadsheet_id = '1csJofpAG3oD204rxkVoYzI5WGjvrMeWsaq2oe2J7hig'
    sheet_number = '3'

    url = 'http://gsx2json.com/api?id=%s&sheet=%s&q=%s&columns=false&integers=false' % (spreadsheet_id, sheet_number, query)

    try:
        response = request.urlopen(url)
        out = json.load(response)
        print(out)
        from pprint import pprint
        pprint(out['rows'])

    except error.HTTPError as e:
        print('HTTPError = ' + str(e.code))
    except error.URLError as e:
        print('URLError = ' + str(e.reason))
    except error.HTTPException as e:
        print('HTTPException')
    except Exception:
        import traceback
        print('generic exception: ' + traceback.format_exc())

    # with request.urlopen(url) as response:
        # html = response.read()




if __name__ == '__main__':
    main()