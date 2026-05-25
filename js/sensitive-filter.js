(function() {
  'use strict';

  window.containsSensitiveWord = function(text) {
    if (typeof text !== 'string' || !text) return null;
    var lower = text.toLowerCase();
    var words = window.SENSITIVE_WORDS;
    if (!words || !words.length) return null;
    for (var i = 0; i < words.length; i++) {
      if (lower.indexOf(words[i].toLowerCase()) !== -1) {
        return words[i];
      }
    }
    return null;
  };

})();
