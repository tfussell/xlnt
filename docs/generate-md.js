var xpath = require('xpath');
var dom = require('xmldom').DOMParser;
var fs = require('fs');
var _ = require('underscore');
var async = require('async');

function generate(directory, primary_callback) {
  var modules = [];
  return fs.readFile(directory + 'index.xml', 'utf8', (err, xml) => {
    if (err) throw err;
    var doc = new dom().parseFromString(xml);
    var nodes = xpath.select('/doxygenindex/compound[@kind="dir"]', doc);
    for (var i = 0; i < nodes.length; i++) {
      var refid = nodes[i].getAttribute('refid');
      var name = nodes[i].firstChild.firstChild.toString();
      modules.push({name: name, refid: refid});
    }
    const root_module = _.find(modules, (module) => {
      return _.last(module.name.split('/')) == 'xlnt';
    });
    const root_name = root_module.name;
    modules = _.filter(modules, (module) => {
      return module.name.indexOf(root_name) >= 0 && module.name !== root_name;
    });
    return async.mapSeries(modules, (module, module_callback) => {
      return fs.readFile(directory + module.refid + '.xml', 'utf8', (err, xml) => {
	if (err) throw err;
	var doc = new dom().parseFromString(xml);
	var nodes = xpath.select('/doxygen/compounddef/innerfile', doc);
	module.source_files = [];
	for (var i = 0; i < nodes.length; i++) {
          var refid = nodes[i].getAttribute('refid');
          var name = nodes[i].firstChild.toString();
          module.source_files.push({name: name, refid: refid});
	}
	return async.map(module.source_files, (source_file, source_file_callback) => {
	  return fs.readFile(directory + source_file.refid + '.xml', 'utf8', (err, xml) => {
	    if (err) throw err;
	    var doc = new dom().parseFromString(xml);
	    var nodes = xpath.select('/doxygen/compounddef/innerclass', doc);
	    module.class_files = [];
	    for (var i = 0; i < nodes.length; i++) {
              var refid = nodes[i].getAttribute('refid');
              var name = nodes[i].firstChild.toString();
              module.class_files.push({name: name, refid: refid});
	    }
	    return async.map(module.class_files, (class_file, class_file_callback) => {
	      return fs.readFile(directory + class_file.refid + '.xml', 'utf8', (err, xml) => {
		if (err) throw err;
		var doc = new dom().parseFromString(xml);
		var nodes = xpath.select('/doxygen/compounddef/sectiondef/memberdef[@prot="public"]', doc);
		var members = [];
		for (var i = 0; i < nodes.length; i++) {
		  var member_data = {id: nodes[i].getAttribute('id')};
		  var child = nodes[i].firstChild;
		  while (child != nodes[i].lastChild) {
		    if (child.textContent.trim()) {
		      member_data[child.nodeName] = child.textContent.trim();
		    }
		    child = child.nextSibling;
		  }
		  members.push(member_data);
		}
		class_file.members = members;
		return class_file_callback(null, class_file);
	      });
	    }, (err, data) => {
	      return source_file_callback(null, _.flatten(data, true));
	    });
	  });
	}, (err, data) => {
	  module.classes = _.flatten(data, true);
	  return module_callback(null, module);
	});
      });
    }, (err, data) => {
      return primary_callback(null, _.flatten(data, true));
    });
  });
}

generate('doxyxml/', function(err, data) {
  console.log('# API Reference');
  for (var i = 0; i < data.length; i++) {
    var module = data[i];
    var module_name = _.last(module.name.split('/'));
    console.log('##', module_name.charAt(0).toUpperCase() + module_name.slice(1), 'Module');
    for (var j = 0; j < module.classes.length; j++) {
      var class_ = module.classes[j];
      console.log('###', _.last(class_.name.split('::')));
      for (var k = 0; k < class_.members.length; k++) {
	var member = class_.members[k];
	console.log('####', '```' + member.definition + member.argsstring + '```');
	if (member.briefdescription) {
	  console.log(member.briefdescription);
	}
      }
    }
  }
});
