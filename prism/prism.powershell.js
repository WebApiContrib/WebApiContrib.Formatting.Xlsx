// With thanks to Brian Marsh: http://blog.briankmarsh.com/prismjs-powershell-syntax-highlighting/
Prism.languages.powershell = {  
  // This comment regex is ugly because prism.js replaces "<" with "&lt;" behind the scenes for some reason
  'comment': /(\&lt\;#[\w\W]*?#>)|(\#.*)/g,
  'string': /(\@\"[\w\W]*?\"\@)|((\'|\")[\w\W]*?(\'|\"))/g,
  'keyword': /\b(switch|if|else|while|do|for|return|function|new|try|throw|catch|finally|break|exit|begin|process|end)(?![-\S])?\b/ig,
  'boolean': /(\$true|\$false)/g,

  // This is for PowerShell Actions, leveraging the theme's pre-defined color scheme for attr-value
  'attr-value': /(add|get|read|test|start|new|set|write|output|where)-\S*/ig,

  // This is for PowerShell Variables, leveraging the theme's pre-defined color scheme for symbol
  'symbol': /(?!(\$true|\$false))(\$[a-z|A-Z|0-9|_|-]*)\b/g,
  'number': /\b-?(0x[\dA-Fa-f]+|\d*\.?\d+([Ee]-?\d+)?)\b/g,
  'operator': /[-+]{1,2}|!|&lt;=?|>=?|={1,3}|(&amp;){1,2}|\|?\||\?|\*|\/|\~|\^|\%|-or|-and|-lt|-le|-gt|-ge|-match|-like/g,
  'ignore': /&(lt|gt|amp);/gi,
  'punctuation': /[{}[\];(),.:]/g
};
