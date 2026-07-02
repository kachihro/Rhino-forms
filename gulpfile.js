'use strict';

const build = require('@microsoft/sp-build-web');
const fs = require('fs');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Sync version from package.json → config/package-solution.json before every build.
// Developer bumps package.json only; the 4-part SPFx version (X.Y.Z.0) is derived automatically.
const versionSync = build.subTask('version-sync', function(gulp, buildOptions, done) {
  const pkg = JSON.parse(fs.readFileSync('./package.json', 'utf8'));
  const [major, minor, patch] = pkg.version.split('.');
  const spfxVersion = `${major}.${minor}.${patch}.0`;

  const solutionPath = './config/package-solution.json';
  const solution = JSON.parse(fs.readFileSync(solutionPath, 'utf8'));
  if (solution.solution.version !== spfxVersion) {
    solution.solution.version = spfxVersion;
    fs.writeFileSync(solutionPath, JSON.stringify(solution, null, 2) + '\n');
    console.log(`[version-sync] solution version → ${spfxVersion}`);
  }
  done();
});

build.rig.addPreBuildTask(versionSync);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
