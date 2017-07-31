var attachments = {
    wfhCardDateInput: (toEmail,accessToken) => {
      return {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: {
            "type": "AdaptiveCard",
            "version": "0.5",
            "body": [
              {
                  "type": "Container",
                  "items": [
                      {
                        "type": "TextBlock",
                        "text": "Send a work from home calendar notification",
                        "weight": "bolder",
                        "size": "large"
                      },
                //       {
                // "type": "ColumnSet",
                // "columns": [
                //   {
                //     "type": "Column",
                //     "size": "stretch",
                //     "items": [
                      {
                        "type": "TextBlock",
                        "text": "From: ",
                        "weight": "bolder"
                      },
                      {
                        "type": "Input.Date",
                        "id": "startdate",
                        "separation": "none"
                      },
                  //   ]
                  // },
                  // {
                  //     "type": "Column",
                  //     "size": "stretch",
                  //     "items": [
                      {
                        "type": "TextBlock",
                        "text": "To: ",
                        "weight": "bolder"
                      },
                      {
                        "type": "Input.Date",
                        "id": "enddate",
                        "separation": "none"
                      },
              //       ]
              //     }
              //   ]
              // },
              // {
              //     "type": "ColumnSet",
              //     "columns": [
              //       {
              //           "type": "Column",
              //           "items":[
                            {
                        "type": "TextBlock",
                        "weight": "bolder",
                        "text": "Subject: "
                      },
                      {
                        "type": "Input.Text",
                        "id": "tbsubject",
                        "value": "Alyssa WFH",
                        "placeholder": "Subject of email (e.g. Emily WFH)",
                        "isMultiline": false,
                        "separation": "none",
                        "isRequired": true
                      },
                      {
                        "type": "TextBlock",
                        "text": "Description:",
                        "weight": "bolder"
                      },
                      {
                        "type": "Input.Text",
                        "id": "tbdescription",
                        "placeholder": "Enter additional details - this will be your email body",
                        "separation": "none",
                        "isRequired": false,
                        "isMultiline": true
                      }
                        ]
                  //   }   
                  //  ]
              // }
              
            // ]
            
              }],"actions": [
                    {
                        'type': 'Action.Submit',
                        'title': 'Send WFH Email',
                        'data': {
                            'type': 'sendWFH'
                        }
                    }
                  ]
          }
      }
    },

          wfhCardDateTB: (dateRange, toEmail, accessToken) => {
            return {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: {
                  "type": "AdaptiveCard",
                  "version": "0.5",
                  "body": [
                    {
                        "type": "Container",
                        "items": [
                            {
                              "type": "TextBlock",
                              "text": "Send a work from home calendar notification",
                              "weight": "bolder",
                              "size": "large"
                            },
                            {
                              "type": "TextBlock",
                              "text": "Dates: " + dateRange.start + " to " + dateRange.end,
                              "weight": "bolder"
                            },
                    {
                        "type": "ColumnSet",
                        "columns": [
                          {
                              "type": "Column",
                              "items":[
                                  {
                              "type": "TextBlock",
                              "weight": "bolder",
                              "text": "Subject: "
                            },
                            {
                              "type": "Input.Text",
                              "id": "tbsubject",
                              "value": "Alyssa WFH",
                              "placeholder": "Subject of email (e.g. Emily WFH)",
                              "isMultiline": false,
                              "separation": "none",
                              "isRequired": true
                            },
                            {
                              "type": "TextBlock",
                              "text": "Description:",
                              "weight": "bolder"
                            },
                            {
                              "type": "Input.Text",
                              "id": "tbdescription",
                              "placeholder": "Enter additional details - this will be your email body",
                              "separation": "none",
                              "isRequired": false,
                              "isMultiline": true
                            }
                              ]
                          }   
                        ]
                    }
                    
                  ]
                  
                    }],"actions": [
                    {
                        'type': 'Action.Submit',
                        'title': 'Send WFH Email',
                        'data': {
                            'type': 'sendWFH'
                        }
                    }
                  ]
                }
      }
    }
}
    
function constructResponse (dateRange, toEmail, subject, description) {
           console.log(dateRange)
           return { 
            "Subject": subject,
            "Body": {
            "ContentType": "HTML",
            "Content": description
            },
            "Start": {
                "DateTime": dateRange.start.toISOString(),
                "TimeZone": "Asia/Singapore"
            },
            "End": {
                "DateTime": dateRange.end.toISOString(),
                "TimeZone": "Asia/Singapore"
            },
            "IsAllDay": true,
            "ResponseRequested": false,
            "ShowAs": "Free", // https://msdn.microsoft.com/en-us/office/office365/api/complex-types-for-mail-contacts-calendar#FreeBusyStatus
            "IsReminderOn": false,
            "Attendees": [
                {
                    "EmailAddress": {
                    "Address": toEmail,
                    "Name": "My team"
                    },
                    "Type": "Required" // Single instance
                }
            ]
          }
}

module.exports = attachments;