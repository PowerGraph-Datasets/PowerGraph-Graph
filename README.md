# PowerGraph-graph
# Benchmarking GNN datasets for PowerGrids- graph-level tasks

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

To reproduce the results presented in the paper, download the following compressed data from [here](https://figshare.com/articles/dataset/PowerGraph/22820534?file=46619158) (~2.7GB, when uncompressed):

```bash
wget -O data.tar.gz "https://figshare.com/ndownloader/files/46619158"
tar -xf data.tar.gz
```

Each dataset folder contains the following files:

- `blist.mat`: branch list also called edge order or edge index
- `of_bi.mat`: binary classification labels ($DNS=0$ or  $DNS\neq0$) 
- `of_reg.mat`: regression labels ($DNS$)
- `of_mc.mat`: multi-class labels (see Table 3 in 'PowerGraph: A power grid benchmark dataset for graph neural networks')
- `Bf.mat`: node feature matrix  (Net active power at bus $P_{net}$, Net apparent power at bus $S_{net}$, Voltage magnitude $V$)
- `Ef.mat`: edge feature matrix (Active power flow $P_{i,j}$, Reactive power flow $Q_{i,j}$, Line reactance $X_{i,j}$, Line rating $lr_{i,j}$)
- `exp.mat`: groundtruth explanation (boolean vector assigning value 1 to edges that have undergone the cascading failure)

## Dataset description

| Dataset    |     Name     | Description                    |
| ---------- | :----------: | ------------------------------ |
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

**--model_name**: transformer / gin / gat/ gcn

**--datatype**: binary / multiclass / regression

**--dataset_name**: uk / ieee24 / ieee39 / ieee118


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
| ├──ieee118
```

Remove the for loop in train_gnn.py if running for a specific **--hidden_dim** and **num_layers**.

The models will be saved in **model** directory


## License

This work is licensed under a CC BY 4.0 license.
