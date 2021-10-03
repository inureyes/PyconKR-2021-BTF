/* eslint-disable no-undef */
import * as ai from "backend.ai-client/backend.ai-client-es6.js";

/* global clearInterval, console, setInterval */

async function connectToManager() {
  if ('baiclient' in globalThis && globalThis.baiclient.ready == true) {
    return Promise.resolve(true);
  } 
  const api_key = "AKIAIOSFODNN7EXAMPLE";
  const secret_key = "wJalrXUtnFEMI/K7MDENG/bPxRfiCYEXAMPLEKEY";
  const api_endpoint = "https://127.0.0.1:8091";
  const clientConfig = new ai.backend.ClientConfig(
    api_key, 
    secret_key,
    api_endpoint 
  );
  globalThis.baiclient = new ai.backend.Client(
    clientConfig,
    `Excel Adapter`,
  );
  globalThis.baiclient.ready = false;
  globalThis.baiclient.get_manager_version().then((response) => {
    globalThis.baiclient.ready = true;
    return Promise.resolve(true);
  }).catch((e) =>{
    return Promise.resolve(false);
  });
}

function errorAsString(e) {
  let result = '';
  for (let key in e) {
    if (e.hasOwnProperty(key)) { 
      result = result + " " + key + "," + e[key] + "|";
    }
  }
  return result;
}
/**
 * Adds two numbers.
 * @customfunction
 * @param first First number 
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second + 5;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Test connection between Excel and Cluster
 * @customfunction TEST_CONNECTION
 * @param invocation Custom function handler
 */
export function testConnection(invocation: CustomFunctions.StreamingInvocation<string>): void { 
  connectToManager().then((response) => {
    invocation.setResult("Succeed");
  }).catch((e) =>{
    let result = errorAsString(e);
    invocation.setResult(result);
  });
}

/**
 * Run super-basic Python Hello World
 * @customfunction hello_python_world
 * @param invocation Custom function handler
 */
 export function hello_python_world(invocation: CustomFunctions.StreamingInvocation<string>): void { 
  connectToManager().then(() => {
    invocation.setResult("Creating...");
    let resources = {"cpu": 1, 
      "mem": "1g",
      "domain": "default",
      "group_name": "default",
      "scaling_group": "default",
      "cluster_mode": "single-node",
      "cluster_size": 1,
      "mounts": [],
      "env": {},
      "resource_opts": {},
      "maxWaitSeconds": 15};
    return globalThis.baiclient.createIfNotExists("cr.backend.ai/stable/python:3.8-ubuntu18.04", "HelloWorld", resources);
  }).then(response => {
    invocation.setResult(`My session is created: ${response.sessionId}`);
    return globalThis.baiclient.execute(response.sessionId, 1, "query", "print('hello python world')", {});
  }).then(response => {
    if (response.result.exitCode === 0) {
      invocation.setResult(response.result.console[0][1]);
    }
  }).catch(err => {
    switch (err.type) {
      case ai.backend.Client.ERR_SERVER:
        invocation.setResult(`Session creation failed: ${err.message}`);
        break;
      default:
        invocation.setResult(`request/response failed: ${err.message}`);
    }
  });
  invocation.onCanceled = () => {
  };  
}


/**
 * Run TensorFlow environment 
 * @customfunction RUN_TRAIN
 * @param invocation Custom function handler
 */
 export function run_tf_environment(invocation: CustomFunctions.StreamingInvocation<string>): void { 
  connectToManager().then(() => {
    invocation.setResult("Creating...");
    let resources = {"cpu": 1, 
      "mem": "6g",
      "domain": "default",
      "group_name": "default",
      "scaling_group": "default",
      "cluster_mode": "single-node",
      "cluster_size": 1,
      "mounts": [],
      "env": {},
      "resource_opts": {},
      "maxWaitSeconds": 15};
    return globalThis.baiclient.createIfNotExists("cr.backend.ai/testing/python-tensorflow:2.5-py38-cuda11.3", "TF", resources);
  }).then(response => {
    invocation.setResult(`My session is created: ${response.sessionId}`);
    return globalThis.baiclient.execute(response.sessionId, 1, "query", "print('hello python world')", {});
  }).then(response => {
    if (response.result.exitCode === 0) {
      invocation.setResult(response.result.console[0][1]);
    }
  }).catch(err => {
    switch (err.type) {
      case ai.backend.Client.ERR_SERVER:
        invocation.setResult(`Session creation failed: ${err.message}`);
        break;
      default:
        invocation.setResult(`request/response failed: ${err.message}`);
    }
  });
  invocation.onCanceled = () => {
  };  
}

/**
 * Run TensorFlow environment test code 
 * @customfunction TRAIN_TEST
 * @param {number[][]} values Multiple ranges of values.
 * @param invocation Custom function handler
 */
 export function run_tf_test_code(values, invocation: CustomFunctions.StreamingInvocation<string>): void { 
  try {
    let result = "";
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        result = result + "," + values[i][j].toString();
      }
    }
    invocation.setResult(result);
  } catch(err) {
  }
  invocation.onCanceled = () => {
  };  
}

/**
 * Train IRIS flower dataset.
 * @customfunction IRIS_TRAIN
 * @param {number[][]} values IRIS data
* @param invocation Custom function handler
 */
 export function run_iris_training(values, invocation: CustomFunctions.StreamingInvocation<string>): void { 
  try {
    let data = JSON.stringify(values);
    invocation.setResult(data);
    connectToManager().then(() => {
      invocation.setResult("Creating...");
      let resources = {"cpu": 1, 
        "mem": "6g",
        "mounts": ['code'],
        "env": {},
        "resource_opts": {},
        "maxWaitSeconds": 30};
      return globalThis.baiclient.createIfNotExists("cr.backend.ai/testing/python-tensorflow:2.5-py38-cuda11.3", "TF_IRIS", resources);
    }).then(response => {
      invocation.setResult(`My session is created: ${response.sessionId}`);
      return globalThis.baiclient.execute("TF_IRIS", 1, "query", 
      `import subprocess;print(subprocess.check_output("cd /home/work/code;python receive_iris_data.py ${data}", shell=True))`, {});
    }).then(async response => {
      if (response.result.exitCode === 0) {
        invocation.setResult(`Training...`);
        let query = await globalThis.baiclient.execute("TF_IRIS", 1, "query", 
        `import subprocess;print(subprocess.check_output("cd /home/work/code;python train_iris_model.py", shell=True))`, {});
        while (query.result.status != 'finished') {
          query = await globalThis.baiclient.execute("TF_IRIS", 1, "query", ``, {});
        }
        return query;
      } else {
        invocation.setResult(`Failed to send data.`);
      }
    }).then(response => {
      if (response.result.exitCode === 0) {
        invocation.setResult(`Training completed.`);
      }
    }).catch(err => {
      switch (err.type) {
        case ai.backend.Client.ERR_SERVER:
          invocation.setResult(`Session creation failed: ${err.message}`);
          break;
        default:
          invocation.setResult(`request/response failed: ${err.message}`);
      }
    });
  } catch(err) {
  }
  invocation.onCanceled = () => {
  };  
}

/**
 * Inferencing IRIS flower class prediction model.
 * @customfunction IRIS
 * @param {number[][]} values IRIS data
 * @param invocation Custom function handler
*/
export function run_iris_inference(values, invocation: CustomFunctions.StreamingInvocation<string>): void { 
  try {
    let data = JSON.stringify(values);
    invocation.setResult(data);
    connectToManager().then(() => {
      invocation.setResult("Creating...");
      let resources = {"cpu": 1, 
        "mem": "6g",
        "mounts": ['code'],
        "env": {},
        "resource_opts": {},
        "maxWaitSeconds": 30};
      return globalThis.baiclient.createIfNotExists("cr.backend.ai/testing/python-tensorflow:2.5-py38-cuda11.3", "TF_IRIS", resources);
    }).then(async response => {
      invocation.setResult(`My session is created: ${response.sessionId}`);
      invocation.setResult(`Inferencing...`);
      let run_id = Math.floor(Math.random() * 50000).toString();
      //globalThis.baiclient.requestTimeout = 15000;
      let query = await globalThis.baiclient.execute("TF_IRIS", run_id, "query", 
      `import subprocess;print(subprocess.check_output("cd /home/work/code;python inference_iris_model.py '${data}'", shell=True).decode().strip())`, {});
      while (query.result.status != 'finished') {
        query = await globalThis.baiclient.execute("TF_IRIS", run_id, "query", ``, {});
      }
      return query;
    }).then(response => {
      if (response.result.exitCode === 0) {
        invocation.setResult(`Inference completed.`);
        ///let result = errorAsString(response.result);
        //if(response.result.console && response.result.console.stdout) {
        //  invocation.setResult(response.result.stdout[0]);
        //}
        let results = response.result.console[0][1].split('\n');
        if (results.length > 1) {
          invocation.setResult(response.result.console[0][1].replace('\n',','));
        } else {
          invocation.setResult(results[0]);
        }
        //return results;
      } else {
        let result = errorAsString(response.result);
        invocation.setResult(result);
      }
    }).catch(err => {
      switch (err.type) {
        case ai.backend.Client.ERR_SERVER:
          invocation.setResult(`Session creation failed: ${err.message}`);
          break;
        default:
          invocation.setResult(`request/response failed: ${err.message}`);
      }
    });
  } catch(err) {
  }
  invocation.onCanceled = () => {
  };
}

/**
 * Train Yahoo stock market dataset.
 * @customfunction YAHOO_STOCK_TRAIN
 * @param {number[][]} values Yahoo stock market data
 * @param invocation Custom function handler
 */
export function run_stock_training(values, invocation: CustomFunctions.StreamingInvocation<string>): void { 
  try {
    let data = JSON.stringify(values);
    invocation.setResult(data);
    connectToManager().then(() => {
      invocation.setResult("Creating...");
      let resources = {"cpu": 4, 
        "mem": "6g",
        "mounts": ['code'],
        "env": {},
        "resource_opts": {},
        "maxWaitSeconds": 30};
      return globalThis.baiclient.createIfNotExists("cr.backend.ai/testing/python-tensorflow:2.5-py38-cuda11.3", "TF_STOCK", resources);
    }).then(response => {
      invocation.setResult(`My session is created: ${response.sessionId}`);
      return globalThis.baiclient.execute("TF_STOCK", 1, "query", 
      `import subprocess;print(subprocess.check_output("cd /home/work/code;python receive_stock_data.py ${data}", shell=True))`, {});
    }).then(async response => {
      if (response.result.exitCode === 0) {
        invocation.setResult(`Training...`);
        let query = await globalThis.baiclient.execute("TF_STOCK", 1, "query", 
        `import subprocess;print(subprocess.check_output("cd /home/work/code;python train_stock_model.py", shell=True))`, {});
        while (query.result.status != 'finished') {
          query = await globalThis.baiclient.execute("TF_STOCK", 1, "query", ``, {});
        }
        return query;
      } else {
        invocation.setResult(`Failed to send data.`);
      }
    }).then(response => {
      if (response.result.exitCode === 0) {
        invocation.setResult(`Training completed.`);
      }
    }).catch(err => {
      switch (err.type) {
        case ai.backend.Client.ERR_SERVER:
          invocation.setResult(`Session creation failed: ${err.message}`);
          break;
        default:
          invocation.setResult(`request/response failed: ${err.message}`);
      }
    });
  } catch(err) {
  }
  invocation.onCanceled = () => {
  };  
}

/**
 * Inferencing Yahoo stock market dataset.
 * @customfunction YAHOO_STOCK
 * @param {number[][]} values Yahoo stock market data
 * @param invocation Custom function handler
 */
export function run_stock_inference(values, invocation: CustomFunctions.StreamingInvocation<string>): void { 
  try {
    let data = JSON.stringify(values);
    invocation.setResult(data);
    connectToManager().then(() => {
      invocation.setResult("Creating...");
      let resources = {"cpu": 4, 
        "mem": "6g",
        "mounts": ['code'],
        "env": {},
        "resource_opts": {},
        "maxWaitSeconds": 30};
      return globalThis.baiclient.createIfNotExists("cr.backend.ai/testing/python-tensorflow:2.5-py38-cuda11.3", "TF_STOCK", resources);
    }).then(async response => {
      invocation.setResult(`My session is created: ${response.sessionId}`);
      invocation.setResult(`Inferencing...`);
      let run_id = Math.floor(Math.random() * 50000).toString();
      //globalThis.baiclient.requestTimeout = 15000;
      let query = await globalThis.baiclient.execute("TF_STOCK", run_id, "query", 
      `import subprocess;print(subprocess.check_output("cd /home/work/code;python inference_stock_model.py '${data}'", shell=True).decode().strip())`, {});
      while (query.result.status != 'finished') {
        query = await globalThis.baiclient.execute("TF_STOCK", run_id, "query", ``, {});
      }
      return query;
    }).then(response => {
      if (response.result.exitCode === 0) {
        invocation.setResult(`Inference completed.`);
        ///let result = errorAsString(response.result);
        //if(response.result.console && response.result.console.stdout) {
        //  invocation.setResult(response.result.stdout[0]);
        //}
        let results = response.result.console[0][1].split('\n');
        if (results.length > 1) {
          invocation.setResult(response.result.console[0][1].replace('\n',','));
        } else {
          invocation.setResult(results[0]);
        }
        //return results;
      } else {
        let result = errorAsString(response.result);
        invocation.setResult(result);
      }
    }).catch(err => {
      switch (err.type) {
        case ai.backend.Client.ERR_SERVER:
          invocation.setResult(`Session creation failed: ${err.message}`);
          break;
        default:
          invocation.setResult(`request/response failed: ${err.message}`);
      }
    });
  } catch(err) {
  }
  invocation.onCanceled = () => {
  };
}

/**
 * Inferencing Yahoo stock market dataset.
 * @customfunction YAHOO_STOCK_PREDICT
 * @param {number[][]} values Yahoo stock market data
 * @return {any[][]} Predicted value
 */
 export async function run_stock_inference_multi(values) { 
  try {
    let data = JSON.stringify(values);
    await connectToManager();
    let resources = {"cpu": 1, 
      "mem": "6g",
      "mounts": ['code'],
      "env": {},
      "resource_opts": {},
      "maxWaitSeconds": 30};
    await globalThis.baiclient.createIfNotExists("cr.backend.ai/testing/python-tensorflow:2.5-py38-cuda11.3", "TF_STOCK", resources);

    let run_id = Math.floor(Math.random() * 50000).toString();
    //globalThis.baiclient.requestTimeout = 15000;
    let query = await globalThis.baiclient.execute("TF_STOCK", run_id, "query", 
    `import subprocess;print(subprocess.check_output("cd /home/work/code;python inference_stock_model.py '${data}'", shell=True).decode().strip())`, {});
    while (query.result.status != 'finished') {
      query = await globalThis.baiclient.execute("TF_STOCK", run_id, "query", ``, {});
    }
    let response = query;
    if (response.result.exitCode === 0) {
      return JSON.parse(response.result.console[0][1]);
    } else {
      let result = errorAsString(response.result);
      return result; //invocation.setResult(result);
    }
  } catch (err) {
    switch (err.type) {
      case ai.backend.Client.ERR_SERVER:
        return [`Session creation failed: ${err.message}`];
      default:
        return [`request/response failed: ${err.message}`];
    }
  };
}