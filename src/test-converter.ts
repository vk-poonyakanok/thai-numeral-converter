import { convertText } from './converter';

function testConverter() {
  const testCases = [
    { input: "123", smart: true, expected: "๑๒๓" },
    { input: "spin9", smart: true, expected: "spin9" },
    { input: "9arm", smart: true, expected: "9arm" },
    { input: "hello 123 world", smart: true, expected: "hello ๑๒๓ world" },
    { input: "www.site123.com", smart: true, expected: "www.site123.com" },
    { input: "123", smart: false, expected: "๑๒๓" },
    { input: "spin9", smart: false, expected: "spin๙" },
  ];

  testCases.forEach(({ input, smart, expected }, index) => {
    const result = convertText(input, smart);
    if (result === expected) {
      console.log(`Test ${index + 1} PASSED: "${input}" (smart=${smart}) -> "${result}"`);
    } else {
      console.error(`Test ${index + 1} FAILED: "${input}" (smart=${smart}) -> expected "${expected}", but got "${result}"`);
    }
  });
}

testConverter();
