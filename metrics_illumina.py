import argparse
import pathlib
import os
import re
import datetime


from interop import py_interop_run_metrics, py_interop_run, py_interop_summary
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots


def read_run(path_run, date):
    try:
        run_metrics = py_interop_run_metrics.run_metrics()
        valid_to_load = py_interop_run.uchar_vector(py_interop_run.MetricCount, 0)
        py_interop_run_metrics.list_summary_metrics_to_load(valid_to_load)
        run_folder = run_metrics.read(f'{path_run}', valid_to_load)
        summary = py_interop_summary.run_summary()
        py_interop_summary.summarize_run_metrics(run_metrics, summary)

        dens = round(float((summary.at(0).at(0)).density().mean())/1000,2)
        pf = round(float((summary.at(0).at(0)).percent_pf().mean()),2)
        q30 = round(float(summary.total_summary().percent_gt_q30()),2)
        date = datetime.datetime.strptime(run_metrics.run_info().date(), "%y%m%d")
        phiX = round(float(summary.total_summary().percent_aligned()),2)
        e = None
    except Exception as error:
        date = datetime.datetime.strptime(date, "%y%m%d")
        dens = None
        q30 = None
        pf = None
        phiX = None
        e = error

    print(f'{path_run} : {e}')

    return {
        "Date": date,
        "Density": dens,
        "%Q30": q30,
        "%PF": pf,
        "%aligned": phiX
    }

def create_graph(df, updates=[], density_min=1000, density_max=1400, q30_min=[60, 75, 80]):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.update_yaxes(fixedrange=True)
    fig.update_yaxes(secondary_y=False, range = [0,100])
    fig.update_yaxes(secondary_y=True, range = [0,max(df["Density"]*2)])

    fig.add_trace(
        go.Scatter(x=df["Date"], y=df["Density"], name="Density"),
        secondary_y=True
    )
    fig.add_trace(
        go.Scatter(x=df["Date"], y=df["%Q30"], name="%Q30"),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(x=df["Date"], y=df["%PF"], name="%PF"),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(x=df["Date"], y=df["%aligned"], name="%aligned"),
        secondary_y=False,
    )

    for i in q30_min:
        fig.add_hrect(y0=i, y1=100, line_width=0, fillcolor="green", opacity=0.2, row=1, col=1, layer="below", secondary_y=False)
    fig.add_hrect(y0=density_min, y1=density_max, line_width=0, fillcolor="green", opacity=0.2, row=1, col=1, layer="below", secondary_y=True)

    for i in updates:
        fig.add_vrect(
            x0=i,
            x1=max(df["Date"]),
            fillcolor="blue", opacity=0.2,
            layer="below", line_width=0,
        )
    fig.show()


def main(path, sequencer_id=False, updates=[], density_min=1000, density_max=1400, q30_min=[60, 75, 80]):
    all_runs = list()
    # all_runs = dict()
    for run in os.listdir(f'{path}'):
        if args.sequencer_id:
            regexp = f"^([0-9]{{6}})_({sequencer_id})_([0-9]{{4}})_([0-9]{{9}}-)?([A-Z0-9]{{5,10}})$"
        else:
            regexp = "^([0-9]{6})_([MNB]{1,2}[0-9]{5,6})_([0-9]{4})_([0-9]{9}-)?([A-Z0-9]{5,10})$"
        m = re.match(regexp, run)
        if m:
            metrics = read_run(f'{path}/{run}', m[1])
            all_runs.append(metrics)

    df = pd.DataFrame(all_runs).sort_values(by=['Date'])
    create_graph(df, updates, density_min, density_max, q30_min)




if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generate Metrics Graph for Sequencing Data')
    parser.add_argument('--path', type=pathlib.Path, help='Path where sequencer write runs', required=True)
    parser.add_argument('--sequencer-id', type=pathlib.Path, help='Sequencer ID (used to find runs from this sequencer)')
    parser.add_argument('--threshold-density-min', type=int, help='Threshold for low density', default=1000)
    parser.add_argument('--threshold-density-max', type=int, help='Threshold for high density', default=1400)
    parser.add_argument('--threshold-q30-min', metavar='N', type=int, nargs='+', default=[60, 75, 80], help='Threshold(s) for minimum Q30 pct (default: 60, 75, 80)')
    parser.add_argument('--updates',
        type=lambda s: datetime.datetime.strptime(s, '%Y-%m-%d'), nargs='+', default=[], help='Date(s) of update(s) (format: YYYY-MM-DD ex: 2024-01-30)')

    args = parser.parse_args()

    main(args.path, args.sequencer_id, args.updates, args.threshold_density_min, args.threshold_density_max, args.threshold_q30_min)
