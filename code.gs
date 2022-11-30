/*To do

Create slideshow rather than copy and rename? This way I can leave instructions in the notes easily, but I'm sure a note can be added via the API.
Menu from sheets. A sidebar could read the headings in row 1 and ask people whether they should be included in the shape, or how they should affect formatting
Offer landscape or portrait at start

*/

//Get the ui as a global variable
var ui = SpreadsheetApp.getUi();

function onOpen() {

  ui.createMenu('Seating Plan Maker')
      .addItem('Make slideshow', 'item1')
      .addToUi();
}

function item1() {
  getClassData();
}


function getClassData() {

  //Create an array to store the required data about each student with which to update the presentation
  var reqs=[];
  
  
  //Get name of active sheet to use for group name
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  group = sheet.getSheetName();

  //Get data range for active sheet
  var range = sheet.getDataRange();  
  var numCols = range.getNumColumns();
  var numRows = range.getNumRows();
  var newRange = sheet.getRange(2, 1, numRows-1, numCols);
  values = newRange.getValues();
  Logger.log(values);
  
  //Iterate through data, adding to reqs
  //Different colours / line types depending on data
  
    for (i in values) {
    var row = values[i];
    var name = row[0];  // First column
    var gender = row[1];  
    var PP = row[2];// Second column
    var FSM = row[3];
    var SEN = row[4];
    var BTL = row[5];
    var MedNeeds = row[6];
    var AddNeeds = row[7];
    var KS2EN = row[8];
    var KS2MA = row[9];
    var PA = row[10];
    
    //Default colours for Fill and Outline
    
    var linestyle = "SOLID";  
     
    var Fred = 0.6;
    var Fgreen= 0.7;
    var Fblue = 1;
    
    var Ored = 0;
    var Ogreen= 0;
    var Oblue = 0;  
      
    if(gender=="F")
    { 
      var Fred = 1;
      var Fgreen= 0.8;
      var Fblue = 0.8;
    };
      
    if(PP=="Y")
    { 
      var Ored = 1;
      var Ogreen= 0.8;
      var Oblue = 0;
    };
      
    if(SEN)
    { 
      var linestyle = "DASH";
    };
    
    var objectID = "abcdefg"+Math.floor(Math.random() * 10000);
      
    reqs.push(
     {
      "createShape": {
        "objectId": objectID,
        "shapeType": "TEXT_BOX",
        "elementProperties": {         
          "pageObjectId": "p",
          "size": {
            "width": {
              "magnitude": 60,
              "unit": "PT"
            },
            "height": {
              "magnitude": 60,
              "unit": "PT"
            }
          },
          "transform": {
            "scaleX": 1,
            "scaleY": 1,
            "translateX": i*30,
            "translateY": i*25,
            "unit": "PT"
          }
        }
      }
    },
    
    {
      "updateShapeProperties": {
        "objectId": objectID,
        "fields": "shapeBackgroundFill.solidFill.color,outline", //This is a field mask, which tells the API which fields are going to be updated.
        "shapeProperties": {
          "shapeBackgroundFill": {
            "solidFill": {
              "color": {
                "rgbColor": {
                   
                  "red": Fred,
                  "green": Fgreen,
                  "blue": Fblue,
               
                }
              }
            }
          },
          "outline": {
            "dashStyle":linestyle,
            "weight":{
              "magnitude": 2,
              "unit": "PT",
            },
            "outlineFill":{
            "solidFill": {
              "color": {
                "rgbColor": {
                   
                  "red": Ored,
                  "green": Ogreen,
                  "blue": Oblue,
                } 
                }
              }
            
            
            }
          }
        }
      }
    },
  
    {
      "insertText": {
        "objectId": objectID,
        "text": name,
        "insertionIndex": 0
      }
    },
      {
      "updateTextStyle": {
        "objectId": objectID,
        "fields": "fontFamily,fontSize",
        "textRange": {
          "type": "ALL"
        },
        "style": {
          "fontFamily": "Ubuntu",
          "fontSize": {
            "magnitude": 9,
            "unit": "PT"
          }
        }
      }
    }
    
  );
      
    };
  
  //Now add title textbox
  reqs.push(
     {
      "createShape": {
        "objectId": "mainTitle",
        "shapeType": "TEXT_BOX",
        "elementProperties": {         
          "pageObjectId": "p",
          "size": {
            "width": {
              "magnitude": 150,
              "unit": "PT"
            },
            "height": {
              "magnitude": 60,
              "unit": "PT"
            }
          },
          "transform": {
            "scaleX": 1,
            "scaleY": 1,
            "translateX": 400,
            "translateY": 20,
            "unit": "PT"
          }
        }
      }
    },
    
    {
      "updateShapeProperties": {
        "objectId": "mainTitle",
        "fields": "shapeBackgroundFill.solidFill.color,outline", //This is a field mask, which tells the API which fields are going to be updated.
        "shapeProperties": {
          "shapeBackgroundFill": {
            "solidFill": {
              "color": {
                "rgbColor": {
                   
                  "red": 1,
                  "green": 1,
                  "blue": 1,
               
                }
              }
            }
          },
          "outline": {
            "dashStyle":linestyle,
            "weight":{
              "magnitude": 2,
              "unit": "PT",
            },
            "outlineFill":{
            "solidFill": {
              "color": {
                "rgbColor": {
                   
                  "red": 0,
                  "green": 0,
                  "blue": 0,
                } 
                }
              }
            
            
            }
          }
        }
      }
    },
  
    {
      "insertText": {
        "objectId": "mainTitle",
        "text": group,
        "insertionIndex": 0
      }
    },
      {
      "updateTextStyle": {
        "objectId": "mainTitle",
        "fields": "fontFamily,fontSize",
        "textRange": {
          "type": "ALL"
        },
        "style": {
          "fontFamily": "Ubuntu",
          "fontSize": {
            "magnitude": 24,
            "unit": "PT"
          }
        }
      }
    }
    
  );
  Logger.log(reqs);
  
 
  // Port of Slides API demo by Wesley Chun to Google Apps Script
// Source: http://wescpy.blogspot.co.uk/2016/11/using-google-slides-api-with-python.html


  /*
  from apiclient import discovery
  from httplib2 import Http
  from oauth2client import file, client, tools
  */
  
  /* DIFF
  No comand line imports but some setup required. 
  1. Resources > Libraires and add library id 1-8n9YfGU1IBDmagna_1xZRHdB3c2jOuFdUrBmUDy64ITRfyhQoXH5lHc H/T Spencer Easton https://plus.google.com/u/0/+SpencerEastonCCS/posts/gK1jmbFH5kT
     [Also change the identifier in the Libraries dialog to SLIDES to make it easy to compare to Wesley's code
  2. Resources > Developer Console Project click on the project link and enable Slides API 
  */

  var TMPLFILE = 'Seating Plan' //File to be copied and renamed
  
  /*
  SCOPES = (
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/presentations',
    )
    store = file.Storage('storage.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
        creds = tools.run_flow(flow, store)
    HTTP = creds.authorize(Http())
    DRIVE  = discovery.build('drive',  'v3', http=HTTP)
    SLIDES = discovery.build('slides', 'v1', http=HTTP)
  */
  
  /* DIFF
  Tip by Romain Vialard that if you use DriveApp you can use the project token
  Running the function also automatically triggers any authorization required
  https://plus.google.com/u/0/+MartinHawksey/posts/aPEdvRiFV9Y
  */
  SLIDES.setTokenService(function(){return ScriptApp.getOAuthToken()});
  
  /*
  rsp = DRIVE.files().list(q="name='%s'" % TMPLFILE).execute().get('files')[0]
  DATA = {'name': 'Google Slides API template DEMO'}
  print('** Copying template %r as %r' % (rsp['name'], DATA['name']))
  DECK_ID = DRIVE.files().copy(body=DATA, fileId=rsp['id']).execute().get('id')
  */
  
  /* DIFF
  As Google Apps Script has built-in DriveApp service we can use this to get our template 
  and copy it
  */
  Logger.log('** Copying template **');
  //Create a new slideshow using name of sheet
  var DECK_ID = DriveApp.getFilesByName(TMPLFILE).next().makeCopy().setName("Seating Plan - "+group).getId();

  
  //Populate the slideshow with the student shapes
  SLIDES.presentationsBatchUpdate(DECK_ID, {'requests': reqs});
  
  Logger.log('DONE'); 
}
