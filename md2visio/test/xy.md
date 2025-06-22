```mermaid
---
config:
    xyChart:
        width: 900
        height: 600
        chartOrientation: vertical
        xAxis:
          showLabel: true
        yAxis:
          showTitle: false
    themeVariables:
        xyChart:
            titleColor: "#ff0000"
---
xychart-beta 
    title "Sales Revenue"
    x-axis X [jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec, d1, d2, d3, d4, d5]
    y-axis Y 3000 --> 12000
    bar  [5000, 6000, 7500, 8200, 9500, 10500, 11000, 10200, 9200, 8500, 7000, 6000, 3000, 3000, 3000, 3000, 3000]
    line [5000, 6000, 7500, 8200, 9500, 10500, 11000, 10200, 9200, 8500, 7000, 6000, 3000, 3000, 3000, 3000, 3000]
```

```mermaid
---
config:
    xyChart:
        width: 900
        height: 900
        chartOrientation: horizontal
        xAxis:
          showLabel: true
        yAxis:
          showTitle: false
    themeVariables:
        xyChart:
            titleColor: "#ff0000"
---
xychart-beta 
    title "Sales Revenue"
    x-axis X 2-->10
    y-axis Y 1 --> 100
    bar  [10, 20, 30,40,50,60,70,80,100]
    line [10, 20, 30,40,50,60,70,80]
```

