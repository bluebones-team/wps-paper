import { sel, doc, $ } from './js/util.js'
import UI from './js/ribbonUI.js'
import app from './js/ribbon.js'
import { Collection, Range } from './js/decorator.js';
Object.prototype.set.call(window, { sel, doc, $, UI, app, Collection, Range, });