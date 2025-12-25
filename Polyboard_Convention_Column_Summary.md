# Polyboard Convention Excel File - Column Summary

## Overview
The `Polyboard_convention.xlsx` file defines conventions for polyboard panels used in Polyboard cabinet design. All panels are 4-sided panels with 2 faces.

## File Structure
- **Total Columns:** 9
- **Total Rows:** 26 (panel types)

## Column Definitions

### Column 1: Component
- **Name:** `Component`
- **Description:** The type of panel used in Polyboard cabinet
- **Notes:** All panels are 4-sided panels with 2 faces

### Column 2: Face_1
- **Name:** `Face_1`
- **Description:** Description of Face 1 for the panel type specified in Column 1
- **Purpose:** Defines the surface/material properties of the first face of the panel

### Column 3: Face_2
- **Name:** `Face_2`
- **Description:** Description of Face 2 for the panel type specified in Column 1
- **Purpose:** Defines the surface/material properties of the second face of the panel

### Column 4: Edge_0
- **Name:** `Edge_0`
- **Description:** Edging code used when the panel has **0 edge banding** (no PVC edge banding on any of the 4 sides)
- **Edge Configuration:** 0 sides edgebanded, 4 sides no PVC edgebanding

### Column 5: Edge_1
- **Name:** `Edge_1`
- **Description:** Edging code used when the panel has **1 edge banding** (1 side edgebanded, 3 sides no PVC edgebanding)
- **Edge Configuration:** 1 side edgebanded, 3 sides no PVC edgebanding

### Column 6: Edge_2_no_connect
- **Name:** `Edge_2_no_connect`
- **Description:** Edging code used when the panel has **2 edge bandings** where the 2 edgebanded sides are **NOT adjacent** to each other
- **Edge Configuration:** 2 sides edgebanded (opposite sides), 2 sides no PVC edgebanding
- **Note:** In a standard 4-sided panel, this means the 2 edgebanded sides are opposite to each other

### Column 7: Edge_2_connect
- **Name:** `Edge_2_connect`
- **Description:** Edging code used when the panel has **2 edge bandings** where the 2 edgebanded sides are **adjacent** to each other
- **Edge Configuration:** 2 sides edgebanded (adjacent sides), 2 sides no PVC edgebanding

### Column 8: Edge_3
- **Name:** `Edge_3`
- **Description:** Edging code used when the panel has **3 edge bandings** (3 sides edgebanded, 1 side no PVC edgebanding)
- **Edge Configuration:** 3 sides edgebanded, 1 side no PVC edgebanding

### Column 9: Edge_4
- **Name:** `Edge_4`
- **Description:** Edging code used when the panel has **4 edge bandings** (all 4 sides edgebanded with PVC edgebanding)
- **Edge Configuration:** 4 sides edgebanded with PVC edgebanding

## Edge Banding Summary

| Edge Count | Column Name | Configuration | Notes |
|------------|-------------|---------------|-------|
| 0 | Edge_0 | 0 sides edgebanded, 4 sides no PVC | No edge banding |
| 1 | Edge_1 | 1 side edgebanded, 3 sides no PVC | Single edge banding |
| 2 (non-adjacent) | Edge_2_no_connect | 2 sides edgebanded (opposite), 2 sides no PVC | Opposite sides |
| 2 (adjacent) | Edge_2_connect | 2 sides edgebanded (adjacent), 2 sides no PVC | Adjacent sides |
| 3 | Edge_3 | 3 sides edgebanded, 1 side no PVC | Three edges banded |
| 4 | Edge_4 | 4 sides edgebanded | All edges banded |

## Usage Notes for Code Vibe Project

1. **Panel Structure:** All panels are 4-sided with 2 faces (Face_1 and Face_2)
2. **Edge Banding Logic:** The convention distinguishes between:
   - Number of edges banded (0-4)
   - For 2-edge configurations: whether edges are adjacent or opposite
3. **Data Lookup:** Use Column 1 (Component) as the key to look up corresponding face and edge codes
4. **Edge Code Selection:** Based on the edge banding configuration, select the appropriate column (Edge_0 through Edge_4, with Edge_2 split into no_connect and connect variants)

## File Location
- **Source File:** `Polyboard_convention.xlsx`
- **Total Panel Types:** 26 different panel types defined

---
*Document created for Code Vibe project planning phase*

