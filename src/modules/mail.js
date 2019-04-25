import axios from 'axios';

// Interface

export default ({
  send: ({ from, to, subject, body }) => new Promise((resolve, reject) => {
    if (!from) return reject('Sender is missed');
    if (!to) return reject('Recievers are missed');
    if (!subject) return reject('Subject is missed');
    if (!body) return reject('Body is missed');
    axios({
      url: '/_api/SP.Utilities.Utility.SendEmail',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'content-type': 'application/json;odata=verbose'
      },
      method: 'POST',
      data: JSON.stringify({
        properties: {
          __metadata: { type: 'SP.Utilities.EmailProperties' },
          From: from,
          To: { results: to },
          Subject: subject,
          Body: body
        }
      })
    }).then(resolve).catch(reject)
  })
})