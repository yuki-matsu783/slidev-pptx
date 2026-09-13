---
theme: default
title: 受入テスト用デッキ（PPT 部品）
colorSchema: light
addons:
  - slidev-addon-pptx
---

## PPT 部品

<PptText :x="60" :y="120" :w="400" fill="#eef" :radius="8" :padding="12" name="lead">

**実測配置**が既定で、`x` `y` を書けば座標指定になります。

</PptText>

<PptShape type="rightArrow" :x="500" :y="120" :w="120" :h="40" fill="#3b82f6" line="none" name="arrow">次へ</PptShape>

<PptShape type="line" :x="60" :y="200" :w="400" :h="0" :line="{ color: '#999', width: 2, tail: 'arrow' }" />

<PptImage src="/bg.png" :x="640" :y="120" :w="280" :h="200" fit="cover" alt="背景" />

<PptTable :x="60" :y="260" :rows="[['名前','用途'],['default','通常'],['cover','表紙']]" />

<PptText :x="500" :y="260" :w="300" export="image">

画像への置き換えを指定した部品

</PptText>

<PptText :x="500" :y="360" :w="300">

<PptShape type="rect" :x="0" :y="0" :w="10" :h="10" />

部品の入れ子（W-NESTED-PPT）

| 表 | は無視 |
|---|---|
| a | b |

</PptText>
