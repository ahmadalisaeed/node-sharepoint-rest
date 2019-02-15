/*
 * decaffeinate suggestions:
 * DS001: Remove Babel/TypeScript constructor workaround
 * DS102: Remove unnecessary code created because of implicit returns
 * DS206: Consider reworking classes to avoid initClass
 * Full docs: https://github.com/decaffeinate/decaffeinate/blob/master/docs/suggestions.md
 */
const SuperClass   = require('./SuperClass');
const UserProfiles = require('./UserProfiles');
const CustomLists  = require('./CustomLists');

class SharePoint extends SuperClass {
  static initClass() {
    this.include(UserProfiles);
    this.include(CustomLists);
  }

  constructor(settings){
    {
      // Hack: trick Babel/TypeScript into allowing this before super.
      if (false) { super(); }
      let thisFn = (() => { return this; }).toString();
      let thisName = thisFn.match(/return (?:_assertThisInitialized\()*(\w+)\)*;/)[1];
      eval(`${thisName} = this;`);
    }
    this.settings = settings;
    UserProfiles.call(this);
    CustomLists.call(this);

    if (!this.settings) {
      throw new Error("settings object is required for instance creation");
    } else {
      if (!this.settings.strictSSL) {
        this.settings.strictSSL = false;
      }

      this.request  = require('request');

      this.user = this.settings.username || undefined;
      this.pass = this.settings.password || undefined;
      this.url  = this.settings.url      || undefined;

      if ((typeof this.url === "undefined") || (typeof this.user === "undefined") || (typeof this.pass === "undefined")) {
        throw new Error("settings object requires username, password, and url for instance creation");
      }

      this.setSiteUrl(this.url);
    }
  }

  log(msg){
    return console.log(msg);
  }

  setSiteUrl(url){
    this.url = url;
    this.log(`setting site url to: ${this.url}`);
    return this;
  }

  getContext(app, cb){
    const processRequest = function(err, res, body){
      if (!body || !JSON.parse(body).d) {
        return console.log(`no list of title: ${app}`);
      } else {
        return cb(err, JSON.parse(body).d.GetContextWebInformation.FormDigestValue);
      }
    };

    const config = {
      headers : {
        Accept: "application/json;odata=verbose"
      },
      strictSSL: this.settings.strictSSL,
      url     : `${this.url}/${app}/_api/contextinfo`
    };

    this.request.post(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }
}
SharePoint.initClass();

module.exports = SharePoint;