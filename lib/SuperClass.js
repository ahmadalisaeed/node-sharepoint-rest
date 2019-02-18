/*
 * decaffeinate suggestions:
 * DS101: Remove unnecessary use of Array.from
 * DS102: Remove unnecessary code created because of implicit returns
 * Full docs: https://github.com/decaffeinate/decaffeinate/blob/master/docs/suggestions.md
 */
const moduleKeywords = ['included', 'extended'];

class SuperClass {
  static include(obj) {
    if (!obj) { throw('include(obj) requires obj'); }
    for (let key in obj.prototype) {
      const value = obj.prototype[key];
      if (!Array.from(moduleKeywords).includes(key)) {
        this.prototype[key] = value;
      }
    }

    const { included } = obj;
    if (included) { included.apply(this); }
    return this;
  }
}

export default SuperClass;