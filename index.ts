let globRange: number[] = [];
let specNumber: number;
async function selectRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    // range.format.fill.color = "yellow";
    // range.load("address");
    range.load();
    await context.sync();
    globRange = range.values.flat(2).filter(Number);
  });
}
async function selectSpecNumber() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    // range.format.fill.color = "yellow";
    // range.load("address");
    range.load();
    await context.sync();
    specNumber = range.values.flat(2).filter(Number)[0];
  });
}
function calculate() {
  subsetSum(globRange, specNumber, []);
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

function subsetSum(numbers, target, partial) {
  var s, n, remaining;

  partial = partial || [];

  // sum partial
  s = partial.reduce(function(a, b) {
    return a + b;
  }, 0);

  // check if the partial sum is equals to target
  if (s === target) {
    console.log("find " + partial.join("+") + " = " + target);
  }

  if (s >= target) {
    return; // if we reach the number why bother to continue
  }

  for (var i = 0; i < numbers.length; i++) {
    n = numbers[i];
    remaining = numbers.slice(i + 1);
    subsetSum(remaining, target, partial.concat([n]));
  }
}

document.getElementById("select-range")!.addEventListener("click", () => tryCatch(selectRange));
document.getElementById("spec-number")!.addEventListener("click", () => tryCatch(selectSpecNumber));
document.getElementById("calculate")!.addEventListener("click", () => {
  subsetSum(globRange, specNumber, []);
});
