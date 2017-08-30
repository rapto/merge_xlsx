# merge_xlsx
Python facilities for generating excel files from xlsx templates, keeping all format, images, charts, etc.

Example usage:
```python
from merge_xlsx import merge
subs = dict(A18='Hello, world')
for i in range(2,16):
    subs ['B%s' % i] = i*i
    subs ['A%s' % i] = dt.datetime(2017,8,i)
merge('demo.xlsx', 'salida.xlsx', **subs)
```
Template:
![Worksheet used as template](../images/images/demo.png)

Result:
![Result Worksheet (merged with data)](../images/images/salida.png)

