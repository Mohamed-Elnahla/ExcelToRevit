# ExcelToRevit - BIM Modeler Assistant Plugin

![License](https://img.shields.io/badge/License-MENCRL%20v1.0-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Autodesk%20Revit-orange.svg)
![Language](https://img.shields.io/badge/Language-C%23-green.svg)

This Revit plugin is part of a comprehensive research framework for **BIM-based Schedule Generation and Optimization Using Genetic Algorithms**. It serves as the BIM Modeler Assistant component, automating the creation and placement of structural elements in Autodesk Revit from Excel data.

> 🔬 **Research Context:** This plugin is a component of the research paper: [BIM-based schedule generation and optimization using genetic algorithms](https://www.sciencedirect.com/science/article/abs/pii/S0926580524002127?via%3Dihub) published in *Automation in Construction* journal.

> 📦 **Complete Framework:** For the full research implementation including schedule optimization, genetic algorithms, and 5D simulation, visit the main repository: [P6-Automation](https://github.com/Mohamed-Elnahla/P6-Automation)

---

## Plugin Features

### 🏗️ **Family Creation from Excel**

The plugin automatically creates Revit families from Excel files for:

1. **Levels** - Generate building levels with elevations
2. **Structural Elements:**
   - Columns
   - Beams
   - Walls
   - Slabs
   - Isolated Footings (RC & PC)
   - Slab Footings
   - Piles
   - Title Blocks

*Excel templates for family creation are included in the plugin resources.*

### 📍 **Element Placement from Excel**

Automatically places structural elements at specified coordinates based on Excel files:

1. **Columns** - Place at grid intersections
2. **Beams** - Position along structural grids
3. **Walls** - Create along specified paths
4. **Isolated Footings** - Place at column bases
5. **Slabs** - Generate floor systems
6. **Piles** - Position pile foundations

*Excel templates for element placement are included in the plugin resources.*

---

## Research Integration

This plugin is integral to the BIM-based automation framework described in our research, specifically:

- **Step 1:** Automated BIM model creation using this ExcelToRevit plugin
- **Step 2:** BIM data extraction for schedule generation
- **Step 3:** Schedule optimization using Genetic Algorithms
- **Step 4:** 5D simulation and Business Intelligence dashboard integration

### Framework Architecture

```
Excel Data → [ExcelToRevit Plugin] → BIM Model → Schedule Generation → GA Optimization → 5D Simulation
```

---

## Installation

1. Download the latest release from the [Releases](../../releases) page
2. Copy the `.addin` and `.dll` files to your Revit Add-ins folder:
   - `%APPDATA%\Autodesk\Revit\Addins\[REVIT_VERSION]\`
3. Restart Autodesk Revit
4. The plugin will appear in the Revit ribbon

---

## Usage

1. **Family Creation:** Use the Excel templates in `Resources/Excel/Families/` to define your structural elements
2. **Element Placement:** Use the Excel templates in `Resources/Excel/Location/` to specify coordinates
3. **Level Creation:** Use `Resources/Excel/Levels.xlsx` to define building levels
4. Access the plugin tools through the Revit ribbon interface

---

## Research Citation

If you use this plugin in your research or find it helpful, please cite our paper:

```bibtex
@article{WEFKI2024105476,
title = {BIM-based schedule generation and optimization using genetic algorithms},
journal = {Automation in Construction},
volume = {164},
year = {2024},
issn = {0926-5805},
doi = {https://doi.org/10.1016/j.autcon.2024.105476},
url = {https://www.sciencedirect.com/science/article/pii/S0926580524002127},
author = {Hossam Wefki and Mohamed Elnahla and Emad Elbeltagi},
keywords = {Building information modeling (BIM), Automation, Genetic algorithm, Automatic Scheduling, Construction, Optimization, Integrated project delivery (IPD), 5D simulation}
}
```

---

## Related Repositories

- **[P6-Automation](https://github.com/Mohamed-Elnahla/P6-Automation)** - Complete research framework including:
  - BIM Data Extraction Plugin
  - Schedule Generation Scripts
  - Genetic Algorithm Optimization
  - 5D Simulation Tools
  - Business Intelligence Dashboard Integration

---

## Developer

**Mohamed Elnahla**
- Email: eng.mohamed.elnahla@gmail.com
- Research focus: BIM Automation, Construction Optimization, Genetic Algorithms

---

## License

This project is licensed under the **Mohamed Elnahla Non-Commercial Research License (MENCRL) v1.0**. 

- ✅ **Permitted:** Academic research, education, non-profit scientific research
- ❌ **Prohibited:** Commercial use without prior agreement
- 📝 **Required:** Proper attribution in all publications and derivative works

The same license terms apply to the main [P6-Automation](https://github.com/Mohamed-Elnahla/P6-Automation) repository.

For commercial licensing, please contact: eng.mohamed.elnahla@gmail.com

---

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests. For major changes, please open an issue first to discuss the proposed changes.

---

*This plugin contributes to advancing automation in the construction industry through integrated BIM-based workflows and optimization techniques.*
