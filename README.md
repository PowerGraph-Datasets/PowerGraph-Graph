# Benchmarking GNN datasets for PowerGrids

Traditional methods for simulating cascading failures in power grids are computationally expensive and use historical blackout datasets that are scarce and incomplete. To address this limitation, we propose to use machine learning models, specifically Graph Neural Networks (GNNs), to instantly detect cascading failures from the pre-outage state of the system. The lack of publicly available GNN datasets for power grid applications has motivated us to develop a new graph dataset. This dataset, designed for graph-level tasks and ranging in size from small to medium, fills a gap in the OGB taxonomy for graph datasets. It is tailored to the society domain, where no public GNN graph property prediction datasets are currently available. Furthermore, we provide explainability masks. Currently, no real graph dataset for graph-level application can be used to benchmark explainability models.

With the InMemoryDatasets Class, we generate the GNN datasets for the UK, IEEE24, IEEE39, IEEE118 power grids. We use **InMemoryDataset** class of Pytorch Geometric.

## Installation

**Requirements**

- CPU or NVIDIA GPU, Linux, Python 3.7
- PyTorch >= 1.5.0, other packages

Load every additional packages:

```
pip install -r requirements.txt
```

## Prerequisites and data structure

To reproduce the results presented in the paper, download the following compressed data from [here](https://figshare.com/articles/dataset/PowerGraph/22820534) (~1.8GB, when uncompressed):

```bash
wget -O data.tar.gz "https://figshare.com/ndownloader/files/40571123"
tar -xf data.tar.gz
```

Each dataset folder contains the following files:

- `blist.mat`: branch list also called edge order or edge index
- `of_bi.mat`: binary classification labels
- `of_reg.mat`: regression labels
- `of_mc.mat`: multi-class labels
- `Bf.mat`: node feature matrix
- `Ef.mat`: edge feature matrix
- `exp.mat`: groundtruth explanation

## Dataset description

| Dataset    |     Name     | Description                    |
| ---------- | :----------: | ------------------------------ |
| BA-2motifs | `ba_2motifs` | Random BA graph with 2 motifs. |
| IEEE-24    |   `ieee24`   | IEEE-24 (Powergrid dataset)    |
| IEEE-39    |   `ieee39`   | IEEE-39 (Powergrid dataset)    |
| IEEE-118   |  `ieee118`   | IEEE-118 (Powergrid dataset)   |
| UK         |     `uk`     | UK (Powergrid dataset)         |

We have created a graph dataset that models cascading failure events, which are the main cause of blackouts in power grids. To generate a comprehensive dataset for different power grids, we used a physics-based cascading failure model called Cascades. This model simulates how failures propagate in the IEEE24, IEEE39, IEEE118 and UK power grids. The output of the model is the final demand not being served (DNS). Our dataset consists of a large set of power grid states, representing the operating conditions before an outage, and is linked to an initial triggering outage (one or more failed elements). Each power grid state is represented as a graph, with a graph-level label assigned based on the results of the physics-based model. The dataset is designed for various graph-level tasks, such as multi-class classification, binary classification, and regression. Bus and branches are the elements of a power grid, buses include loads and generators which represent the nodes of the graph, while branches include transmission lines and transformers which represent the edges of the graph. We provide three features per node: net active power, net apparent power and voltage magnitude. While the features per edge are four: active power flow, reactive power flow, line reactance and line rating.

## GNN Benchmarking

To test the datasets with different GNN architectures: GCN, GINe, GAT and Transformer, run,

```
python code/train_gnn.py
```

We have the main arguments to control namely

**--model_name**: transformer / gin / gat

**--datatype**: binary / multiclass / regression

**--dataset_name**: uk / ieee24 / ieee39

We do not use GCN model for the PowerGraph datasets since the graphs have edge attributes that carry meaningful information and should be used by the GNN model for prediction.

Make sure you have the dataset as per format. Models will be saved as per format (make sure you have the model folder)

```
.
├── code
├── dataset
│ ├── processed
│ ├── raw
| | ├── \*.mat
├──model
| ├──ieee24
| ├──ieee39
| ├──uk
```

Remove the for loop in train_gnn.py if running for a specific **--hidden_dim** and **num_layers**.

The models will be saved in **model** directory

## GNN Explainability Benchmarking

**Graph Classification Tasks**

| Explainer            | Paper                                                                                                                       |
| :------------------- | :-------------------------------------------------------------------------------------------------------------------------- |
| Occlusion            | [Visualizing and understanding convolutional networks](https://arxiv.org/pdf/1311.2901.pdf)                                 |
| SA                   | [Explainability Techniques for Graph Convolutional Networks](https://arxiv.org/pdf/1905.13686.pdf)                          |
| Grad-CAM             | [Grad-CAM: Visual Explanations from Deep Networks via Gradient-based Localization](https://arxiv.org/pdf/1610.02391.pdf)    |
| Integrated Gradients | [Axiomatic Attribution for Deep Networks](https://arxiv.org/pdf/1703.01365.pdf)                                             |
| GNNExplainer         | [GNNExplainer: Generating Explanations for Graph Neural Networks](https://arxiv.org/pdf/1903.03894.pdf)                     |
| SubgraphX            | [On Explainability of Graph Neural Networks via Subgraph Exploration](https://arxiv.org/pdf/2102.05152.pdf)                 |
| PGM-Explainer        | [PGM-Explainer: Probabilistic Graphical Model Explanations for Graph Neural Networks](https://arxiv.org/pdf/2010.05788.pdf) |

**Run explainability methods**

```bash
python3 code/main.py --dataset_name [dataset-name] --model_name [gnn-model] --explainer_name [explainer-name]
```

We have the main arguments to control namely

**--model_name**: transformer / gin / gat

**--datatype**: multiclass

**--dataset_name**: uk_mc / ieee24_mc / ieee39_mc

**--explainer_name**: random / sa / ig / gradcam / occlusion / basic_gnnexplainer / gnnexplainer / subgraphx / pgmexplainer / pgexplainer / graphcfe

Default args:

**--explained_target**: 0 - Only Category A powergrids are explained (DNS>0 and Cascading failure)

## License

This work is licensed under a CC BY 4.0 license.
