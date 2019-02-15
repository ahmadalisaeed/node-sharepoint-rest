/*
 * decaffeinate suggestions:
 * DS101: Remove unnecessary use of Array.from
 * DS102: Remove unnecessary code created because of implicit returns
 * DS205: Consider reworking code to avoid use of IIFEs
 * DS207: Consider shorter variations of null checks
 * Full docs: https://github.com/decaffeinate/decaffeinate/blob/master/docs/suggestions.md
 */
class CustomLists {
  constructor(){}

  getLists(cb){
    const processRequest = function(err, res, body){
      if (!body || !JSON.parse(body).d) {
        return cb({err});
      } else {
        return cb(err, JSON.parse(body).d.results);
      }
    };

    const config = {
      headers : {
        Accept: "application/json;odata=verbose"
      },
      strictSSL: this.settings.strictSSL,
      url     : `${this.url}/_api/lists`
    };

    this.request.get(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }

  getListItemsByTitle(title, cb){
    return this.getListItemsByTitleWithQuery(title,'',cb);
  }

  getListItemsByTitleWithQuery(title, query, cb){
    const processRequest = function(err, res, body){
      if (!body || !JSON.parse(body).d) {
        return cb(body);
      } else {
        return cb(err, JSON.parse(body).d.results);
      }
    };

    const config = {
      headers: {
        Accept: "application/json;odata=verbose"
      },
      strictSSL: this.settings.strictSSL,
      url: `${this.url}/_api/web/lists/getbytitle('${title}')/items?${query}`
    };

    this.request.get(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }

  getListTypeByTitle(title, cb){
    const processRequest = function(err, res, body){
      if (!body || !JSON.parse(body).d) {
        return cb(`no list of title : ${title}`);

      } else {
        return cb(err, JSON.parse(body).d.ListItemEntityTypeFullName);
      }
    };

    const config = {
      headers : {
        Accept: "application/json;odata=verbose"
      },
      strictSSL: this.settings.strictSSL,
      url     : `${this.url}/_api/web/lists/getbytitle('${title}')?`
    };

    this.request.get(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }

  addAttachmentToListItem(req, cb){
    const processRequest = function(err, res, body){
      if (err) {
        this.log(err);
      }

      const jsonBody = JSON.parse(body);

      if (jsonBody.error && (jsonBody.error.code.indexOf("Microsoft.SharePoint.Client.InvalidClientQueryException") >= 0)) {
        cb("Microsoft.SharePoint.Client.InvalidClientQueryException", null);
      }

      if (jsonBody.error && jsonBody.error.code.indexOf("Microsoft.SharePoint.SPException")) {
        cb("Microsoft.SharePoint.SPException", null);
      }

      if (jsonBody.error && jsonBody.error.code) {
        return cb(jsonBody.error, null);

      } else {
        return cb(err, JSON.parse(body).d);
      }
    };

    const config = {
      headers : {
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": req.context,
        "content-type": "application/json;odata=verbose"
      },
      url     : `${this.url}/_api/web/lists/getbytitle('${req.title}')/items(${req.itemId})/AttachmentFiles/add(Filename='${req.binary.fileName}')`,
      strictSSL: this.settings.strictSSL,
      body: req.data,
      binaryStringRequestBody: true,
      state: "update"
    };

    this.request.post(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }

  addListItemByTitle(title, item, context, cb){
    const processRequest = function(err, res, body){
      const jsonBody = JSON.parse(body);

      if (jsonBody.error && (jsonBody.error.code.indexOf("Microsoft.SharePoint.Client.InvalidClientQueryException") >= 0)) {
        return cb("Microsoft.SharePoint.Client.InvalidClientQueryException", null);

      } else if (jsonBody.error && jsonBody.error.code) {
        return cb(JSON.parse(body).error, null);

      } else {
        return cb(err, JSON.parse(body).d);
      }
    };

    let itemPayload = {
      '__metadata': {
        'type': this.getItemTypeForListName(title)
      }
    };

    itemPayload = this.merge(itemPayload,item);

    const config = {
      headers : {
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": context,
        "content-type": "application/json;odata=verbose"
      },
      url: `${this.url}/_api/web/lists/getbytitle('${title}')/items`,
      strictSSL: this.settings.strictSSL,
      body: JSON.stringify(itemPayload)
    };

    this.request.post(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }

  editListItemByTitle(title, id, item, context, cb){
    const processRequest = function(err, res, body){
      let jsonBody;
      try {
        jsonBody = JSON.parse(body);
      } catch (e) {
        cb();
        return;
      }

      if (jsonBody.error) {
        console.log(jsonBody);
        return cb(jsonBody, null);
      } else {
        return cb(err, JSON.parse(body).d);
      }
    };

    let itemPayload = {
      '__metadata': {
        'type': this.getItemTypeForListName(title)
      }
    };

    itemPayload = this.merge(itemPayload,item);

    const config = {
      headers : {
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": context,
        "content-type": "application/json;odata=verbose",
        "X-HTTP-Method": "MERGE",
        "If-Match": "*"
      },
      url: `${this.url}/_api/web/lists/getbytitle('${title}')/items(${id})`,
      strictSSL: this.settings.strictSSL,
      body: JSON.stringify(itemPayload)
    };

    this.request.post(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }

  deleteListItemByTitle(title, id, context, cb){
    const processRequest = function(err, res, body){
      let jsonBody;
      try {
        jsonBody = JSON.parse(body);
      } catch (e) {
        cb();
        return;
      }

      if (jsonBody.error) {
        console.log(jsonBody);
        return cb(jsonBody, null);
      } else {
        return cb(err, JSON.parse(body));
      }
    };

    const config = {
      headers : {
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": context,
        "content-type": "application/json;odata=verbose",
        "X-HTTP-Method": "DELETE",
        "If-Match": "*"
      },
      url: `${this.url}/_api/web/lists/getbytitle('${title}')/items(${id})`,
      strictSSL: this.settings.strictSSL
    };

    this.request.post(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }

  createList(req, cb){
    if (!req.context) {
      cb({err: "please provide a context"});
      return;
    }

    if (!req.title) {
      cb({err: "please provide a title"});
      return;
    }

    if (!req.description) {
      cb({err: "please provide a description"});
      return;
    }

    const { context }     = req;
    const { title }       = req;
    const { description } = req;

    const processRequest = function(err, res, body){
      const jsonBody = JSON.parse(body);

      if (jsonBody.error && (jsonBody.error.code.indexOf("Microsoft.SharePoint.Client.InvalidClientQueryException") >= 0)) {
        return cb("Microsoft.SharePoint.Client.InvalidClientQueryException", null);

      } else if (jsonBody.error && jsonBody.error.code) {
        return cb(JSON.parse(body).error, null);

      } else {
        return cb(err, JSON.parse(body).d);
      }
    };

    const body = {
      __metadata: {
        type: 'SP.List'
      },
      AllowContentTypes: true,
      BaseTemplate: 100,
      ContentTypesEnabled: true,
      Description: description,
      Title: title
    };

    const config = {
      headers : {
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": context,
        "content-type": "application/json;odata=verbose"
      },
      url: `${this.url}/_api/web/lists`,
      strictSSL: this.settings.strictSSL,
      body: JSON.stringify(body)
    };

    this.request.post(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }

  deleteListByGUID(req, cb){
    if (!req.context) {
      cb({err: "please provide a context"});
      return;
    }

    if (!req.guid) {
      cb({err: "please provide a guid"});
      return;
    }

    if (!cb) {
      cb({err: "please provide a callback"});
      return;
    }

    const { context } = req;
    const { guid }    = req;

    const processRequest = (err, res, body)=> cb(err);

    const config = {
      headers: {
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": context,
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE"
      },
      url: `${this.url}/_api/web/lists(guid'${guid}')`,
      strictSSL: this.settings.strictSSL
    };

    this.request.post(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }

  createColumnForListByGUID(req, cb){
    if (!req.context) {
      cb({err: "please provide a context"});
      return;
    }

    if (!req.title) {
      cb({err: "please provide a title"});
      return;
    }

    if (!req.type) {
      cb({err: "please provide a type"});
      return;
    }

    if (!req.guid) {
      cb({err: "please provide a guid"});
      return;
    }

    const { context } = req;
    const { title }   = req;
    const { type }    = req;
    const { guid }    = req;

    const processRequest = function(err, res, body){
      const jsonBody = JSON.parse(body);
      if (jsonBody.error && jsonBody.error.code) {
        return cb(jsonBody.error, null);

      } else {
        return cb(err, JSON.parse(body).d);
      }
    };

    const body = {
      __metadata: {
        type: 'SP.Field'
      },
      Title: title,
      FieldTypeKind: type,
      Required: 'false',
      EnforceUniqueValues: 'false',
      StaticName: title
    };

    if (type === 3) {
      body.__metadata.RichText = "TRUE";
      body.__metadata.RichTextMode = "FullHtml";
    }

    const config = {
      headers : {
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": context,
        "content-type": "application/json;odata=verbose"
      },
      url: `${this.url}/_api/web/lists(guid'${guid}')/Fields`,
      strictSSL: this.settings.strictSSL,
      body: JSON.stringify(body)
    };

    this.request.post(config, processRequest).auth(this.user, this.pass, true);

    return this;
  }

  getItemTypeForListName(name) {
    return `SP.Data.${name.charAt(0).toUpperCase()}${name.slice(1)}ListItem`;
  }

  merge(...xs) {
    if ((xs != null ? xs.length : undefined) > 0) {
      return this.tap({}, m=> Array.from(xs).map((x) => (() => {
        const result = [];
        for (let k in x) {
          const v = x[k];
          result.push(m[k] = v);
        }
        return result;
      })()) );
    }
  }

  tap(o, fn){ fn(o); return o; }
}

module.exports = CustomLists;
