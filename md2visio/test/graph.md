# Subgraph
```mermaid
---
title: 带文本的节点
---
%%{init: {'theme':'forest'}}%%
 graph LR   
    C
    --> 
    D[[D估算载荷分布]]    
    
    C --> SJ
    F[F遍历激光雷达数据点] --> G{G是否位于A/B段?}
    
    subgraph SJ ["实际扰度
              [计算]" 'calc']
       direction
       G -- 是 --> H[H计算对应位置的载荷 P]
             
       subgraph ad [" "]
        direction TB
       	G x-- LINK-=TEXT --> F
       end       
    end
```

```mermaid
graph LR
   subgraph 结果处理
        G -->|否| D
        G -->|是| H(H:结果分析)
        H --> I(I:可视化结果)
    end    

    subgraph 初始化
        A[A:初始化系统] --> B(B:设置仿真与控制参数)
        B --> C(C:初始化AUV状态和环境参数)
    end
    
    subgraph 主循环
        D{D:主循环} --> E(E:更新状态与时间)
        E --> F(F:计算控制输出)
        F --> G{G:结束循环?}
    end 
    
    C --> D
```

```mermaid
%%{init: {'theme':'neutral'}}%%
graph TB
    subgraph TL [推力控制]
        C --> D[D计算速度误差]
        D --> E[E计算PID控制输出]
        E --> F[F调整发动机转速]
        F --> G[G计算推力]
    end
    subgraph "位置控制"
        K
        C
        I -->|已达到| J[J检查是否到达目标位置]
        subgraph Inner
        	K-->C
        end
    end

    A[A开始] --> B[B设定期望速度和目标位置]
    B --> TL
    G --> H[H更新装备速度]
    H --> I[I检查是否达到期望速度]
    I -->|未达到| C[C读取当前速度和位置]
    J -->|未到达| K[K调整方向和速度]
    J -->|已到达| L[L结束]
```



## Comment

`````mermaid
%%{init: {"flowchart": {"htmlLabels": false}, 'theme':'dark'} }%%
graph RL
	%% this is a comment
    %%
    A%%
    
    %%
    -- Text1 ---
    A%%
    
    D""
    `` x==x E --> A%%
    
    A%% -->|Text2|D""
`````

# Edge and Shape

````mermaid
%%{init: {'theme':'base'}}%%
graph RL
    - -.-x D
    
    A ==o xD
    A-o--xD   
    
    Ax--xD
    Ax--text0
    -->
    A --"text1"
    text2
    --> E 
    
    D --> E{"`E{构建理想挠
    	曲线**模型**}`"}
    
    D[
    D:lonely
    ]
        
    : <-- to --> E ~~~~ F
`````


# AMP

```````mermaid
graph BT
    &A & B --> C
    C-->D
```````

# Style

```mermaid
graph LR
	classDef className fill:#f9f,stroke:#333,stroke-width:4px;
	class A className;
	style A-B fill:#bbf,stroke:#f66,stroke-width:2px,color:#fff,stroke-dasharray: 5 5
	A-B-->B@{shape: rounded, label: 'A: 文件,处理'}
```

# HTML

```mermaid
graph LR
    A[确定目标距离 D] --> B[计算航行距离 D<sub>nav</sub>]
    B --> C{D ≤ D<sub>nav</sub> ?}
    C -->|是| D[计算自毁伤害 H]
    D --> E[执行撞击动作]
    C -->|否| F[继续搜索目标或采取其他行动]
```

