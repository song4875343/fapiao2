import filltable

# 测试修复后的代码
logic = filltable.logic
amount, invoice_no = logic.extract_invoice_info('【风韵出行-24.01元-1个行程】高德打车电子发票.pdf')

print(f'发票号: {invoice_no}')
print(f'金额: {amount}')

if amount == 24.01:
    print('\n✅ 修复成功！金额识别正确！')
else:
    print(f'\n❌ 仍有问题，期望: 24.01, 实际: {amount}')
