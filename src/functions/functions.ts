/* global clearInterval, console, CustomFunctions, setInterval */
/**
 * 年终奖单独计税 VS 年终奖合并计税
 * @customfunction
 * @param yearincome 年累计税前收入
 * @param yearaward 年终奖
 * @param wuxianyijin 年累计五险一金
 * @param special 年累计特殊扣除
 * @param other 年其他累计扣除
 * @param isSeparete 是否单独计税
 * @returns 返回年累计缴税
 */
export function tax(
  yearincome: number,
  yearaward: number,
  wuxianyijin: number,
  special: number,
  other: number,
  isSeparete: boolean
): number {
  if (isSeparete) {
    let yearTax = taxCal(yearincome - 60000 - wuxianyijin - special - other);
    let awardTax = taxCal(yearaward, true);
    return yearTax + awardTax;
  } else {
    return taxCal(yearincome + yearaward - 60000 - wuxianyijin - special - other);
  }
}
function taxCal(taxincome: number, ismonthquick: boolean = false): number {
  let rate: number = 0;
  let quick: number = 0;
  switch (true) {
    case taxincome <= 0:
      rate = 0;
      quick = 0;
      break;
    case taxincome <= 36000:
      rate = 0.03;
      quick = 0;
      break;
    case taxincome <= 144000:
      rate = 0.1;
      quick = 2520;
      break;
    case taxincome <= 300000:
      rate = 0.2;
      quick = 16920;
      break;
    case taxincome <= 420000:
      rate = 0.25;
      quick = 31920;
      break;
    case taxincome <= 660000:
      rate = 0.3;
      quick = 52920;
      break;
    case taxincome <= 960000:
      rate = 0.35;
      quick = 85920;
      break;
    default:
      rate = 0.45;
      quick = 181920;
      break;
  }
  if (ismonthquick) {
    quick = quick / 12;
  }
  console.log(`${taxincome} * ${rate} - ${quick}`);
  return taxincome * rate - quick;
}
/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second + 600;
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
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}
