const spawn = require('child_process').spawn;

module.exports = function() {
  const args = Array.prototype.slice.call(arguments);

  return new Promise((resolve, reject) => {
    const cp = spawn.apply(null, args);
    let stdout = '', stderr = '';

    cp.stdout.on('data', (chunk) => {
      stdout += chunk;
    });

    cp.stderr.on('data', (chunk) => {
      stderr += chunk;
    });

    cp.on('error', reject)
      .on('close', (code) => {
        if (code === 0) { resolve(stdout); } 
        else { reject(stderr); 
        }
      });    
  });
};