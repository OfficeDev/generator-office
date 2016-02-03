'use strict';

module.exports = {
  app: require.resolve('./generators/app'),
  content: require.resolve('./generators/content'),
  mail: require.resolve('./generators/mail'),
  taskpane: require.resolve('./generators/taskpane'),
  commands: require.resolve('./generators/commands')
};
