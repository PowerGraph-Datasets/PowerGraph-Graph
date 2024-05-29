"""
File to train the GNN model.
"""

import os
from pathlib import Path
import torch
import shutil
import warnings
import numpy as np
from torch.optim import Adam
from utils.parser_utils import (
    arg_parse,
    create_args_group,
    fix_random_seed,
    get_data_args,
    get_graph_size_args,
)
import pandas as pd
from utils.io_utils import check_dir
from gendata import get_dataloader, get_dataset
import torch.nn.functional as F
from torch.optim.lr_scheduler import ReduceLROnPlateau
from gnn.model import get_gnnNets
from sklearn.metrics import balanced_accuracy_score, f1_score, r2_score, roc_auc_score, precision_score, recall_score
import json


# Save directory model_name + dataset_name + layers + hidden_dim

class TrainModel(object):
    def __init__(
        self,
        model,
        dataset,
        device,
        seed,
        graph_classification=False,
        graph_regression=True,
        save_dir=None,
        save_name="model",
        **kwargs,
    ):
        self.model = model
        self.dataset = dataset  # train_mask, eval_mask, test_mask
        self.loader = None
        self.device = device
        self.graph_classification = graph_classification
        self.graph_regression = graph_regression
        #self.node_classification = not graph_classification
        self.optimizer = None
        self.save = save_dir is not None
        self.save_dir = save_dir
        self.save_name = save_name
        self.seed = seed
        check_dir(self.save_dir)

        if self.graph_classification or self.graph_regression:
            self.dataloader_params = kwargs.get("dataloader_params")
            self.loader = get_dataloader(dataset, **self.dataloader_params)

    def __loss__(self, logits, labels):
        if self.graph_classification:
            return F.nll_loss(logits, labels)
        elif self.graph_regression:
            return F.mse_loss(logits.squeeze(), labels)

    # Get the loss, apply optimizer, backprop and return the loss

    def _train_batch(self, data, labels):
        logits = self.model(data)
        if self.graph_classification:
            loss = self.__loss__(logits, labels)
            #print(loss)
        elif self.graph_regression:
            loss = self.__loss__(logits, labels)
        else:
            loss = self.__loss__(logits[data.train_mask], labels[data.train_mask])           

        self.optimizer.zero_grad()
        loss.backward()
        #torch.nn.utils.clip_grad_value_(self.model.parameters(), clip_value=2.0)
        self.optimizer.step()
        return loss.item()

    def _eval_batch(self, data, labels, **kwargs):
        self.model.eval()
        logits = self.model(data)
        if self.graph_classification:
            loss = self.__loss__(logits, labels)
            loss = loss.item()
            preds = logits.argmax(-1)
            return loss, preds, logits
        elif self.graph_regression:
            loss = self.__loss__(logits, labels)
            loss = loss.item()
            preds = logits
            return loss, preds
        else:
            mask = kwargs.get("mask")
            if mask is None:
                warnings.warn("The node mask is None")
                mask = torch.ones(labels.shape[0])
            loss = self.__loss__(logits[mask], labels[mask])
            loss = loss.item()
            preds = logits.argmax(-1)

        return loss, preds

    def eval(self):
        self.model.to(self.device)
        self.model.eval()

        if self.graph_classification:
            losses, accs, balanced_accs, f1_scores = [], [], [], []
            for batch in self.loader[0]["eval"]:
                batch = batch.to(self.device)
                loss, batch_preds, logits = self._eval_batch(batch, batch.y)
                losses.append(loss)
                accs.append(batch_preds == batch.y)
                balanced_accs.append(balanced_accuracy_score(batch.y.cpu(), batch_preds.cpu()))
                f1_scores.append(f1_score(batch.y.cpu(), batch_preds.cpu(), average="weighted"))
            eval_loss = torch.tensor(losses).mean().item()
            eval_acc = torch.cat(accs, dim=-1).float().mean().item()
            eval_balanced_acc = np.mean(balanced_accs)
            eval_f1_score = np.mean(f1_scores)
            print(
                f"Test loss: {eval_loss:.4f}, test acc {eval_acc:.4f}, balanced test acc {eval_balanced_acc:.4f}, test f1 score {eval_f1_score:.4f}"
            )
            return eval_loss, eval_acc, eval_balanced_acc, eval_f1_score
        elif self.graph_regression:
            losses, r2scores = [], []
            for batch in self.loader[0]["eval"]:
                batch = batch.to(self.device)
                loss, batch_preds = self._eval_batch(batch, batch.y)
                r2scores.append(r2_score(batch.y.cpu(), batch_preds.cpu()))
                losses.append(loss)
            eval_loss = torch.tensor(losses).mean().item()
            eval_r2score = np.mean(r2scores)
            print(
                f"eval loss: {eval_loss:.6f}, eval r2score {eval_r2score:.6f}"
            )
            return eval_loss, eval_r2score
        else:
            data = self.dataset.data.to(self.device)
            eval_loss, preds = self._eval_batch(data, data.y, mask=data.val_mask)
            eval_acc = (preds == data.y).float().mean().item()
            eval_balanced_acc = balanced_accuracy_score(data.y, preds)
            eval_f1_score = f1_score(data.y, preds, average="weighted")
        return eval_loss, eval_acc, eval_balanced_acc, eval_f1_score

    def test(self):
        state_dict = torch.load(
            os.path.join(self.save_dir, f"{self.save_name}_best.pth")
        )["net"]
        self.model.load_state_dict(state_dict)
        self.model = self.model.to(self.device)
        self.model.eval()
        #cross test
        """
        from dataset import (
            SynGraphDataset,
            PowerGrid
        )
        dataset1 = PowerGrid(
            "C:\\Users\\avarbella\\Documents\\01_GraphGym\\PowerGraph-master\\code\\dataset\\ieee118\\", 'ieee118',
            datatype='binary')
        dataset1.data.x = dataset1.data.x.float()
        dataset1.data.y = dataset1.data.y.squeeze().long()
        #dataloader_params = self.kwargs.get("dataloader_params")
        loader1 = get_dataloader(dataset1, **self.dataloader_params)
        preds_c, balanced_accs_c = [], []
        for batch in loader1[0]["test"]:
            batch = batch.to(self.device)
            loss, batch_preds = self._eval_batch(batch, batch.y)
            preds_c.append(batch_preds)
            balanced_accs_c.append(balanced_accuracy_score(batch.y.cpu(), batch_preds.cpu()))
        test_balanced_acc_c = np.mean(balanced_accs_c)

        print(
            f"balanced test acc {test_balanced_acc_c:.4f}"
        )
        """
        if self.graph_classification:
            losses, preds, targets, accs, balanced_accs = [], [], [], [], []
            for batch in self.loader[0]["test"]:
                batch = batch.to(self.device)
                loss, batch_preds, logits = self._eval_batch(batch, batch.y)
                losses.append(loss)
                preds.append(logits)
                targets.append(batch.y)
                accs.append(batch_preds == batch.y)
                balanced_accs.append(balanced_accuracy_score(batch.y.cpu(), batch_preds.cpu()))

            test_loss = torch.tensor(losses).mean().item()
            preds = torch.vstack(preds)
            targets = torch.cat(targets, dim=-1)
            test_acc = torch.cat(accs, dim=-1).float().mean().item()
            test_balanced_acc = np.mean(balanced_accs)

            print(
                f"Test loss: {test_loss:.4f}, test acc {test_acc:.4f}, balanced test acc {test_balanced_acc:.4f}"
            )
            scores = {
            "test_loss": test_loss,
            "test_acc": test_acc,
            "test_balanced_acc": test_balanced_acc,
            }
            self.save_scores(scores)
            return test_loss, test_acc, test_balanced_acc, preds, targets
        elif self.graph_regression:
            losses, r2scores, preds, targets = [], [], [], []
            for batch in self.loader[0]["test"]:
                batch = batch.to(self.device)
                loss, batch_preds = self._eval_batch(batch, batch.y)
                preds.append(batch_preds.squeeze())
                targets.append(batch.y)
                r2scores.append(r2_score(batch.y.detach().cpu(), batch_preds.detach().cpu()))
                losses.append(loss)
            test_loss = torch.tensor(losses).mean().item()
            test_r2score = np.mean(r2scores)
            preds = torch.cat(preds, dim=-1)
            targets = torch.cat(targets, dim=-1)
            print(
                f"test loss: {test_loss:.6f}, test r2score {test_r2score:.6f}"
            )
            scores = {
            "test_loss": test_loss,
            "test r2score": test_r2score,
            }
            self.save_scores(scores)
            return test_loss, test_r2score, preds, targets
        else:
            data = self.dataset.data.to(self.device)
            test_loss, preds = self._eval_batch(data, data.y, mask=data.test_mask)
            test_acc = (preds == data.y).float().mean().item()
            test_balanced_acc = balanced_accuracy_score(data.y, preds)
            test_f1_score = f1_score(data.y, preds, average="weighted")
            print(
                f"Test loss: {test_loss:.4f}, test acc {test_acc:.4f}, balanced test acc {test_balanced_acc:.4f}, test f1 score {test_f1_score:.4f}"
            )
            scores = {
            "test_loss": test_loss,
            "test_acc": test_acc,
            "test_balanced_acc": test_balanced_acc,
            "test_f1_score": test_f1_score,
            }
            self.save_scores(scores)
            return test_loss, test_acc, test_balanced_acc, test_f1_score, preds

    # Train model
    def train(self, train_params=None, optimizer_params=None):
        if self.graph_classification:
            num_epochs = train_params["num_epochs"]
            num_early_stop = train_params["num_early_stop"]
            #milestones = train_params["milestones"] # needed if using a different LR scheduler
            #gamma = train_params["gamma"]
            
            if optimizer_params is None:
                self.optimizer = Adam(self.model.parameters())
            else:
                self.optimizer = Adam(self.model.parameters(), **optimizer_params)
            
            lr_schedule = ReduceLROnPlateau(
                self.optimizer, mode='max', factor=0.2, patience=10, verbose=True
            )
            
            self.model.to(self.device)
            best_eval_acc = 0.0
            best_eval_loss = 0.0
            early_stop_counter = 0
            for epoch in range(num_epochs):
                is_best = False
                self.model.train()
                if self.graph_classification:
                    losses = []
                    for batch in self.loader[0]["train"]:
                        batch = batch.to(self.device)
                        loss = self._train_batch(batch, batch.y)
                        losses.append(loss)
                    train_loss = torch.FloatTensor(losses).mean().item()
            
                else:
                    data = self.dataset.data.to(self.device)
                    train_loss = self._train_batch(data, data.y)
            
                with torch.no_grad():
                    eval_loss, eval_acc, eval_balanced_acc, eval_f1_score = self.eval()
            
                print(
                    f"Epoch:{epoch}, Training_loss:{train_loss:.4f}, Eval_loss:{eval_loss:.4f}, Eval_acc:{eval_acc:.4f}, Eval_balanced_acc:{eval_balanced_acc:.4f}, Eval_f1_score:{eval_f1_score:.4f}, lr:{self.optimizer.param_groups[0]['lr']}"
                )
                if num_early_stop > 0:
                    if eval_loss <= best_eval_loss:
                        best_eval_loss = eval_loss
                        early_stop_counter = 0
                    else:
                        early_stop_counter += 1
                    if epoch > num_epochs / 2 and early_stop_counter > num_early_stop:
                        break
                if lr_schedule:
                    lr_schedule.step(eval_acc)
            
                if best_eval_acc < eval_acc:
                    is_best = True
                    best_eval_acc = eval_acc
                recording = {"epoch": epoch, "is_best": str(is_best)}
                if self.save:
                    self.save_model(is_best, recording=recording)
        
        elif self.graph_regression:
            num_epochs = train_params["num_epochs"]
            num_early_stop = train_params["num_early_stop"]
            # milestones = train_params["milestones"] # needed if using a different LR scheduler
            # gamma = train_params["gamma"]

            if optimizer_params is None:
                self.optimizer = Adam(self.model.parameters())
            else:
                self.optimizer = Adam(self.model.parameters(), **optimizer_params)

            lr_schedule = ReduceLROnPlateau(
                self.optimizer, mode='min', factor=0.1, patience=10
            )

            self.model.to(self.device)
            best_eval_r2score = -100.0
            best_eval_loss = 0.0
            early_stop_counter = 0
            for epoch in range(num_epochs):
                is_best = False
                self.model.train()
                if self.graph_regression:
                    losses = []
                    for batch in self.loader[0]["train"]:
                        batch = batch.to(self.device)
                        loss = self._train_batch(batch, batch.y)
                        losses.append(loss)
                    train_loss = torch.FloatTensor(losses).mean().item()

                else:
                    data = self.dataset.data.to(self.device)
                    train_loss = self._train_batch(data, data.y)

                with torch.no_grad():
                    eval_loss, eval_r2score = self.eval()

                print(
                    f"Epoch:{epoch}, Training_loss:{train_loss:.4f}, Eval_loss:{eval_loss:.4f}, Eval_r2:{eval_r2score:.4f}, lr:{self.optimizer.param_groups[0]['lr']}"
                )
                if num_early_stop > 0:
                    if eval_loss <= best_eval_loss:
                        best_eval_loss = eval_loss
                        early_stop_counter = 0
                    else:
                        early_stop_counter += 1
                    if epoch > num_epochs / 2 and early_stop_counter > num_early_stop:
                        break
                if lr_schedule:
                    lr_schedule.step(eval_loss)

                if best_eval_r2score < eval_r2score:
                    is_best = True
                    best_eval_r2score = eval_r2score
                recording = {"epoch": epoch, "is_best": str(is_best)}
                if self.save:
                    self.save_model(is_best, recording=recording)

    # Save each latest and best model
    def save_model(self, is_best=False, recording=None):
        self.model.to("cpu")
        state = {"net": self.model.state_dict()}
        for key, value in recording.items():
            state[key] = value
        latest_pth_name = f"{self.save_name}_latest.pth"
        best_pth_name = f"{self.save_name}_best.pth"
        ckpt_path = os.path.join(self.save_dir, latest_pth_name)
        torch.save(state, ckpt_path)
        if is_best:
            print("saving best...")
            shutil.copy(ckpt_path, os.path.join(self.save_dir, best_pth_name))
        self.model.to(self.device)

    # Load the best model
    def load_model(self):
        state_dict = torch.load(
            os.path.join(self.save_dir, f"{self.save_name}_best.pth")
        )["net"]
        self.model.load_state_dict(state_dict)
        self.model.to(self.device)

    def save_scores(self, scores):
        with open(os.path.join(self.save_dir, f"{self.save_name}_scores.json"), "w") as f:
            json.dump(scores, f)

#  Main train function
def train_gnn(args, args_group):
    fix_random_seed(args.seed)
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    print(f"dev {device}")

    dataset_params = args_group["dataset_params"]
    model_params = args_group["model_params"]
    if args.datatype == 'regression':
        model_params['graph_regression'] = "True"
    else:
        model_params['graph_regression'] = "False"

    # changing the dataset path here, load the dataset
    dataset = get_dataset(
        dataset_root=os.path.join(args.data_save_dir, args.dataset_name),
        **dataset_params,
    )
    dataset.data.x = dataset.data.x.float()

    if args.datatype == 'regression':
        args.graph_regression = "True"
        args.graph_classification = "False"
        dataset.data.y = dataset.data.y.squeeze().float()

    else:
        args.graph_regression = "False"
        args.graph_classification = "True"
        dataset.data.y = dataset.data.y.squeeze().long()
    # get dataset args
    args = get_data_args(dataset, args)


    if eval(args.graph_classification) | eval(args.graph_regression):
        dataloader_params = {
            "batch_size": args.batch_size,
            "random_split_flag": eval(args.random_split_flag),
            "data_split_ratio": [args.train_ratio, args.val_ratio, args.test_ratio],
            "seed": args.seed,
        }
    # get model
    model = get_gnnNets(args.num_node_features, args.num_classes, model_params, eval(args.graph_regression))

    # train model
    if eval(args.graph_classification):
        trainer = TrainModel(
            model=model,
            dataset=dataset,
            device=device,
            seed=args.seed,
            graph_classification=eval(args.graph_classification),
            graph_regression=eval(args.graph_regression),
            save_dir=os.path.join(args.model_save_dir, args.dataset_name),
            save_name=f"{args.dataset_name}_{args.model_name}_{args.datatype}_{args.num_layers}l_{args.hidden_dim}h_{args.seed}s",
            dataloader_params=dataloader_params,
        )
    else:
        trainer = TrainModel(
            model=model,
            dataset=dataset,
            device=device,
            seed=args.seed,
            graph_classification=eval(args.graph_classification),
            graph_regression=eval(args.graph_regression),
            save_dir=os.path.join(args.model_save_dir, args.dataset_name),
            save_name=f"{args.dataset_name}_{args.model_name}_{args.datatype}_{args.num_layers}l_{args.hidden_dim}h_{args.seed}s",
            dataloader_params=dataloader_params,
        ) 

    if Path(os.path.join(trainer.save_dir, f"{trainer.save_name}_best.pth")).is_file():
        trainer.load_model()
    else:
        trainer.train(
            train_params=args_group["train_params"],
            optimizer_params=args_group["optimizer_params"],
        )
    # test model
    if eval(args.graph_regression):
        test_loss, test_r2score, preds, targets = trainer.test()
        predicted_values = preds.detach().cpu().numpy() if isinstance(preds, torch.Tensor) else np.array(
            preds)
        target_values = targets.cpu().numpy() if isinstance(targets, torch.Tensor) else np.array(targets)
        data = {'Target Values': target_values, 'Predicted Values': predicted_values}

        # Create a DataFrame for metrics
        metrics = {'Metric': ['MSE loss', 'R2 score'],
                   'Value': [test_loss, test_r2score]}

        # Combine data and metrics DataFrames
        df_data = pd.DataFrame(data)
        df_metrics = pd.DataFrame(metrics)

        # Combine data and metrics DataFrames
        save_dir = os.path.join(args.model_save_dir, args.dataset_name)
        save_name = f"summary{args.dataset_name}_{args.model_name}_{args.datatype}_{args.num_layers}l_{args.hidden_dim}h_{args.seed}s"

        # Create an Excel writer with the specified directory and file name
        output_path = os.path.join(save_dir, f"{save_name}.xlsx")

        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Write data to the sheet named 'Data'
            df_data.to_excel(writer, sheet_name='Data', index=False)

            # Write metrics to the sheet named 'Metrics'
            df_metrics.to_excel(writer, sheet_name='Metrics', index=False)
    else:
        _, test_acc, test_balanced_acc, preds, targets = trainer.test()

        probs = torch.nn.functional.softmax(preds, dim=1)

        # Get predicted class for each sample
        _, predicted_classes = torch.max(probs, 1)

        predicted_probs = probs.detach().cpu().numpy() if isinstance(preds, torch.Tensor) else np.array(
            probs)
        predicted_values = predicted_classes.detach().cpu().numpy() if isinstance(preds, torch.Tensor) else np.array(
            predicted_classes)
        target_values = targets.cpu().numpy() if isinstance(targets, torch.Tensor) else np.array(targets)
        f1 = f1_score(target_values, predicted_values, average='weighted')
        if args.datatype == 'binary':
            roc_auc = roc_auc_score(target_values, predicted_values, average='weighted')
        else:
            roc_auc = roc_auc_score(target_values, predicted_probs, multi_class='ovr', average='weighted')
        precision = precision_score(target_values, predicted_values, average='weighted', zero_division=0.0)
        recall = recall_score(target_values, predicted_values, average='weighted', zero_division=0.0)
        # Create a DataFrame for data
        data = {'Target Values': target_values, 'Predicted Values': predicted_values}

        # Create a DataFrame for metrics
        metrics = {'Metric': ['ACC', 'BAL ACC', 'F1-score', 'ROC AUC', 'Precision', 'Recall'],
                   'Value': [test_acc, test_balanced_acc, f1, roc_auc, precision, recall]}

        # Combine data and metrics DataFrames
        df_data = pd.DataFrame(data)
        df_metrics = pd.DataFrame(metrics)

        # Combine data and metrics DataFrames
        save_dir = os.path.join(args.model_save_dir, args.dataset_name)
        save_name = f"summary{args.dataset_name}_{args.model_name}_{args.datatype}_{args.num_layers}l_{args.hidden_dim}h_{args.seed}s"

        # Create an Excel writer with the specified directory and file name
        output_path = os.path.join(save_dir, f"{save_name}.xlsx")

        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Write data to the sheet named 'Data'
            df_data.to_excel(writer, sheet_name='Data', index=False)

            # Write metrics to the sheet named 'Metrics'
            df_metrics.to_excel(writer, sheet_name='Metrics', index=False)


if __name__ == "__main__":
    parser, args = arg_parse()
    args = get_graph_size_args(args)

    # for loop the training architecture for the number of layers and hidden dimensions
    rnd_seeds = [0, 100, 300, 700, 1000]
    tasks = ['regression', 'binary', 'multiclass']
    powergrids = ['ieee24', 'uk', 'ieee39', 'ieee118']
    models = ['gat', 'gcn', 'gin', 'transformer']
    for powergrid in powergrids:
        args.dataset_name = powergrid
        for task in tasks:
            args.datatype = task
            for model in models:
                args.model_name = model
                for rnd_seed in rnd_seeds:
                    args.seed = rnd_seed
                    fix_random_seed(rnd_seed)
                    args.num_layers = 3
                    args.hidden_dim = 16

                    args_group = create_args_group(parser, args)
                    print(f"Hidden_dim: {args.hidden_dim}, Num_layers: {args.num_layers}, model {args.model_name}, data {args.dataset_name}, task {args.datatype}, rnd_seed {rnd_seed} ")
                    train_gnn(args, args_group)

        print("CHANGE POWERGRID")

    print("END TRAINING OF POWERGRAPH")
