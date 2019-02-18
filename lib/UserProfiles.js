/*
 * decaffeinate suggestions:
 * DS102: Remove unnecessary code created because of implicit returns
 * Full docs: https://github.com/decaffeinate/decaffeinate/blob/master/docs/suggestions.md
 */
class UserProfiles {
  constructor(){}

  getPropertiesForAccountName(accountName, cb){
    const processRequest = function(err, res, body){
      if (!body || !JSON.parse(body).d) {
        return cb(`no account of : ${accountName}`);
      } else {
        return cb(err, JSON.parse(body).d);
      }
    };

    const config = {
      headers: {
        Accept: "application/json;odata=verbose"
      },
      strictSSL: this.settings.strictSSL,
      url: `${this.url}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${accountName}'`
    };

    this.request.get(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }
}

export default UserProfiles;
