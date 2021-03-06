/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */
export function add(first: number, second: number): number {
  return first + second;
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
    console.log(result)
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}


/**
 * Writes a message to console.log().
 * @customfunction STREAMING
 * @param message String to write.
 * @param {CustomFunctions.StreamingInvocation<string[][]>} invocation Uses the invocation parameter present in each cell.
 */
export function streaming(message: string, invocation?: CustomFunctions.StreamingInvocation<string[][]>): void {
  invocation.setResult([['Retrieving...']]);
  try {
    const timeout = setTimeout(() => {
      console.log('render');
    invocation.setResult([["A", "B"],["C", "D"]] as string[][]);
    }, 30000)
    invocation.onCanceled = () => {
      clearTimeout(timeout);
    };
  } catch (e) {
    console.error(e);
  }
  
  
}

/**
 * Writes a message to console.log().
 * @customfunction STREAMING2
 * @param {string[][]} valueOne String to write.
 * @param {string[][]} valueTwo String to write.
 * @param {CustomFunctions.StreamingInvocation<string[][]>} invocation Uses the invocation parameter present in each cell.
 */
 export function streaming2(valueOne: string[][], valueTwo: string[][],invocation?: CustomFunctions.StreamingInvocation<string[][]>): void {
  try {
    const valueOneString = flatten2DArrayToString(valueOne);
    const valueTwoString = flatten2DArrayToString(valueTwo);
   
    invocation.setResult([[valueOneString, valueTwoString]]);
  } catch (e) {
    console.error(e);
  }
  
}

function flatten2DArrayToString(valueList: string[][]): string {
  let stringValue = '';
  valueList.forEach((values: string[]) => {
    values.forEach((value: string) => {
      stringValue = `${stringValue},${value}`
    });
  });
  return stringValue;
}
