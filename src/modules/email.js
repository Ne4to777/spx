import axios from 'axios';

import {
  ACTION_TYPES,
  AbstractBox,
  getArray,
  map,
  arrayInit,
  arrayLast,
  prop,
  getInstance,
  switchCase,
  typeOf,
  identity,
  ifThen,
  isArrayFilled,
  constant
} from './../lib/utility'

import site from './../modules/site';

// Internal

const NAME = 'email';

const EMAIL_RE = /@.+\./;

const getSender = opts => site.user.get({ ...opts, isSP: true, view: 'Email' });
const getRecievers = opts => users => site.user(users).get({ ...opts, isSP: true, view: 'EMail' });

const liftMailType = switchCase(typeOf)({
  object: identity,
  default: _ => ({
    To: ''
  })
})

class Box extends AbstractBox {
  constructor(value = '') {
    super(value);
    this.value = this.isArray
      ? ifThen(isArrayFilled)([
        map(liftMailType),
        constant([liftMailType()])
      ])(value)
      : liftMailType(value);
  }
  getCount() {
    return this.isArray ? this.value.length : 1;
  }
}

// Interface

export default params => {
  const instance = {
    box: getInstance(Box)(params)
  }
  return {
    send: async (opts = {}) => {
      const { isFake, silent, silentInfo, detailed } = opts;
      const result = await instance.box.chain(async element => {
        const { From, To, Subject, Body } = element;
        const recievers = getArray(To);
        if (!recievers.length) throw new Error('Recievers are missed');
        if (!Subject) throw new Error('Subject is missed');
        if (!Body) throw new Error('Body is missed');

        const availableRecievers = [];
        const missedUsers = [];
        map(el => (EMAIL_RE.test(el) ? availableRecievers : missedUsers).push(el))(recievers);
        let recieversEmails = availableRecievers;
        let senderEmail = From;
        if (!From) {
          if (missedUsers.length) {
            const [sender, recievers] = await Promise.all([getSender(), getRecievers(missedUsers)]);
            senderEmail = sender.Email;
            recieversEmails = recieversEmails.concat(recievers.map(prop('EMail')));
          } else {
            senderEmail = (await getSender(opts)).Email;
          }
        } else if (!EMAIL_RE.test(From)) {
          const foundUsers = await getRecievers(opts)(missedUsers.concat(From));
          recieversEmails = recieversEmails.concat(arrayInit(foundUsers).map(prop('EMail')));
          senderEmail = arrayLast(foundUsers).EMail;
        } else if (missedUsers.length) {
          recieversEmails = recieversEmails.concat((await getRecievers(opts)(missedUsers)).map(prop('EMail')));
        }
        if (isFake) return;
        const response = await axios({
          url: '/_api/SP.Utilities.Utility.SendEmail',
          headers: {
            'Accept': 'application/json;odata=verbose',
            'content-type': 'application/json;odata=verbose'
          },
          method: 'POST',
          data: JSON.stringify({
            properties: {
              __metadata: { type: 'SP.Utilities.EmailProperties' },
              From: senderEmail,
              To: { results: recieversEmails },
              Subject,
              Body
            }
          })
        });
        !silent && !silentInfo && detailed && console.log(`Email is sent\nSender: ${senderEmail}\nRecievers: ${recieversEmails.join(', ')}`);
        return response;
      })
      !silent && !silentInfo && console.log(`${
        ACTION_TYPES.send} ${
        instance.box.getCount()} ${
        NAME}(s)`);
      return result
    }
  }
}