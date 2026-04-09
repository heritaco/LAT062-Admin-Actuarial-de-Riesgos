from __future__ import annotations

from collections import defaultdict
from dataclasses import asdict, dataclass
from decimal import Decimal, ROUND_HALF_UP
from itertools import product
from pathlib import Path
import csv
import json
import math
import matplotlib
from zipfile import ZIP_DEFLATED, ZipFile
from typing import Dict, Iterable, List, Tuple
from xml.sax.saxutils import escape

matplotlib.use("Agg")

from matplotlib import pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import PercentFormatter

from .data import (
    BondSpec,
    EXERCISE_BONDS,
    RATING_SPREADS,
    STATES_18,
    TRANSITION_MATRIX_RAW_18,
    TREASURY_CURVE,
)


@dataclass
class BondDistributionRow:
    state: str
    probability: float
    horizon_value: float
    loss_vs_current: float


@dataclass
class BondSummary:
    bond: BondSpec
    current_price: float
    expected_horizon_value: float
    expected_loss: float
    volatility: float
    var_999: float
    distribution: List[BondDistributionRow]


@dataclass
class PortfolioSummary:
    current_value: float
    expected_horizon_value: float
    expected_loss: float
    volatility: float
    var_999: float
    es_999: float
    lower_tail_value_001: float
    support_size_cents: int
    support_size_bruteforce: int
    convolution_matches_bruteforce_cents: bool
    quantile_state_example: Tuple[str, ...]


def round_cents(value: float) -> int:
    return int((Decimal(str(value)) * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def price_bond_today(bond: BondSpec) -> float:
    coupon = bond.face_value * bond.coupon_rate
    rating_spread = RATING_SPREADS[bond.rating]
    total = 0.0
    for t in range(1, bond.maturity_years):
        y = TREASURY_CURVE[t] + rating_spread
        total += coupon / ((1.0 + y) ** t)
    y_final = TREASURY_CURVE[bond.maturity_years] + rating_spread
    total += (coupon + bond.face_value) / ((1.0 + y_final) ** bond.maturity_years)
    return total


def price_bond_at_horizon_given_state(bond: BondSpec, migrated_state: str, horizon_years: int = 1) -> float:
    """Valor al horizonte de 1 año.

    Supuesto estándar CreditMetrics para bono con cupón anual:
    - se recibe el cupón del año 1 en el horizonte;
    - si NO incumple, se revalúa el bono remanente con vencimiento residual de 4 años;
    - si incumple, se usa valor de recuperación sobre el nominal.
    """
    if migrated_state == "D":
        return bond.face_value * bond.recovery_rate

    coupon = bond.face_value * bond.coupon_rate
    remaining_years = bond.maturity_years - horizon_years
    rating_key = migrated_state if migrated_state != "CCC" else "CCC"
    rating_spread = RATING_SPREADS[rating_key]

    # Cupón recibido al final del primer año.
    total = coupon

    # Precio ex-cupón del bono remanente valuado en t=1.
    for u in range(1, remaining_years):
        y = TREASURY_CURVE[u] + rating_spread
        total += coupon / ((1.0 + y) ** u)
    y_final = TREASURY_CURVE[remaining_years] + rating_spread
    total += (coupon + bond.face_value) / ((1.0 + y_final) ** remaining_years)
    return total


def normalized_transition_row(initial_rating: str, nr_policy: str = "renormalize") -> Dict[str, float]:
    raw = TRANSITION_MATRIX_RAW_18["CCC" if initial_rating == "CCC" else initial_rating]
    total = sum(raw)
    if nr_policy == "renormalize":
        probs = [x / total for x in raw]
    elif nr_policy == "keep_current":
        nr_mass = 1.0 - total
        probs = list(raw)
        current_idx = STATES_18.index(initial_rating)
        probs[current_idx] += nr_mass
    else:
        raise ValueError("nr_policy debe ser 'renormalize' o 'keep_current'.")
    return {state: prob for state, prob in zip(STATES_18, probs)}


def single_bond_distribution(bond: BondSpec, nr_policy: str = "renormalize") -> List[BondDistributionRow]:
    probs = normalized_transition_row(bond.rating, nr_policy=nr_policy)
    current = price_bond_today(bond)
    rows = []
    for state in STATES_18:
        value = price_bond_at_horizon_given_state(bond, state)
        rows.append(
            BondDistributionRow(
                state=state,
                probability=probs[state],
                horizon_value=value,
                loss_vs_current=current - value,
            )
        )
    return rows


def summarize_single_bond(bond: BondSpec, nr_policy: str = "renormalize") -> BondSummary:
    dist = single_bond_distribution(bond, nr_policy=nr_policy)
    current = price_bond_today(bond)
    expected = sum(row.horizon_value * row.probability for row in dist)
    variance = sum(((row.horizon_value - expected) ** 2) * row.probability for row in dist)
    losses = sorted((row.loss_vs_current, row.probability) for row in dist)
    var_999 = discrete_quantile(losses, 0.999)
    return BondSummary(
        bond=bond,
        current_price=current,
        expected_horizon_value=expected,
        expected_loss=current - expected,
        volatility=math.sqrt(variance),
        var_999=var_999,
        distribution=dist,
    )


def compress_distribution_to_cents(rows: Iterable[BondDistributionRow]) -> Dict[int, float]:
    pmf: Dict[int, float] = defaultdict(float)
    for row in rows:
        pmf[round_cents(row.horizon_value)] += row.probability
    return dict(sorted(pmf.items()))


def convolve_pmfs(left: Dict[int, float], right: Dict[int, float]) -> Dict[int, float]:
    out: Dict[int, float] = defaultdict(float)
    for value_left, prob_left in left.items():
        for value_right, prob_right in right.items():
            out[value_left + value_right] += prob_left * prob_right
    return dict(sorted(out.items()))


def portfolio_convolution_steps(
    bonds: Iterable[BondSpec], nr_policy: str = "renormalize"
) -> List[Tuple[str, Dict[int, float]]]:
    bonds = list(bonds)
    iterator = iter(bonds)
    first = next(iterator)
    acc = compress_distribution_to_cents(single_bond_distribution(first, nr_policy=nr_policy))
    steps = [(f"convolucion_1_{first.name.lower()}", acc)]
    for step_idx, bond in enumerate(iterator, start=2):
        acc = convolve_pmfs(acc, compress_distribution_to_cents(single_bond_distribution(bond, nr_policy=nr_policy)))
        steps.append((f"convolucion_{step_idx}_bonos", acc))
    return steps


def portfolio_distribution_convolution(bonds: Iterable[BondSpec], nr_policy: str = "renormalize") -> Dict[int, float]:
    return portfolio_convolution_steps(bonds, nr_policy=nr_policy)[-1][1]


def portfolio_distribution_bruteforce(
    bonds: Iterable[BondSpec], nr_policy: str = "renormalize"
) -> Tuple[Dict[int, float], List[Tuple[Tuple[str, ...], float, float, int]]]:
    bonds = list(bonds)
    per_bond = [single_bond_distribution(bond, nr_policy=nr_policy) for bond in bonds]

    pmf: Dict[int, float] = defaultdict(float)
    scenario_rows: List[Tuple[Tuple[str, ...], float, float, int]] = []

    index_space = [range(len(STATES_18)) for _ in bonds]
    for combo_idx in product(*index_space):
        states = tuple(STATES_18[i] for i in combo_idx)
        prob = 1.0
        value = 0.0
        value_cents = 0
        for rows, i in zip(per_bond, combo_idx):
            row = rows[i]
            prob *= row.probability
            value += row.horizon_value
            value_cents += round_cents(row.horizon_value)
        pmf[value_cents] += prob
        scenario_rows.append((states, prob, value, value_cents))
    return dict(sorted(pmf.items())), scenario_rows


def pmf_close(left: Dict[int, float], right: Dict[int, float], tol: float = 1e-15) -> bool:
    keys = set(left) | set(right)
    return all(abs(left.get(k, 0.0) - right.get(k, 0.0)) <= tol for k in keys)


def discrete_quantile(loss_probability_pairs: List[Tuple[float, float]], alpha: float) -> float:
    cumulative = 0.0
    for loss, prob in sorted(loss_probability_pairs, key=lambda x: x[0]):
        cumulative += prob
        if cumulative >= alpha:
            return loss
    return loss_probability_pairs[-1][0]


def expected_shortfall_from_loss_pmf(loss_pmf: Dict[float, float], alpha: float) -> float:
    items = sorted(loss_pmf.items(), key=lambda x: x[0])
    cumulative = 0.0
    tail_mass = 1.0 - alpha
    tail_weighted_sum = 0.0
    for loss, prob in items:
        next_cumulative = cumulative + prob
        if next_cumulative <= alpha:
            cumulative = next_cumulative
            continue
        effective_prob = prob if cumulative >= alpha else next_cumulative - alpha
        tail_weighted_sum += loss * effective_prob
        cumulative = next_cumulative
    return tail_weighted_sum / tail_mass


def portfolio_distribution_rows(pmf_conv: Dict[int, float]) -> List[Tuple[float, float, float]]:
    rows: List[Tuple[float, float, float]] = []
    cumulative = 0.0
    for value_cents, prob in sorted(pmf_conv.items()):
        cumulative += prob
        rows.append((value_cents / 100.0, prob, cumulative))
    return rows


def pmf_comparison_rows(
    pmf_conv: Dict[int, float], pmf_brute: Dict[int, float]
) -> List[Tuple[float, float, float, float, float, float]]:
    rows: List[Tuple[float, float, float, float, float, float]] = []
    cumulative_conv = 0.0
    cumulative_brute = 0.0
    for value_cents in sorted(set(pmf_conv) | set(pmf_brute)):
        prob_conv = pmf_conv.get(value_cents, 0.0)
        prob_brute = pmf_brute.get(value_cents, 0.0)
        cumulative_conv += prob_conv
        cumulative_brute += prob_brute
        rows.append(
            (
                value_cents / 100.0,
                prob_conv,
                prob_brute,
                prob_conv - prob_brute,
                cumulative_conv,
                cumulative_brute,
            )
        )
    return rows


def portfolio_metric_rows(nr_policy: str, portfolio_summary: PortfolioSummary) -> List[Tuple[str, object]]:
    return [
        ("nr_policy", nr_policy),
        ("current_value", portfolio_summary.current_value),
        ("expected_horizon_value", portfolio_summary.expected_horizon_value),
        ("expected_loss", portfolio_summary.expected_loss),
        ("volatility", portfolio_summary.volatility),
        ("var_999", portfolio_summary.var_999),
        ("es_999", portfolio_summary.es_999),
        ("lower_tail_value_001", portfolio_summary.lower_tail_value_001),
        ("support_size_cents", portfolio_summary.support_size_cents),
        ("support_size_bruteforce", portfolio_summary.support_size_bruteforce),
        ("convolution_matches_bruteforce_cents", portfolio_summary.convolution_matches_bruteforce_cents),
        ("quantile_state_example", " | ".join(portfolio_summary.quantile_state_example)),
    ]


def bond_summary_rows(bond_summaries: Iterable[BondSummary]) -> List[Tuple[object, ...]]:
    return [
        (
            summary.bond.name,
            summary.bond.rating,
            summary.current_price,
            summary.expected_horizon_value,
            summary.expected_loss,
            summary.volatility,
            summary.var_999,
        )
        for summary in bond_summaries
    ]


def excel_sheet_title(raw_title: str) -> str:
    invalid_chars = '[]:*?/\\'
    cleaned = "".join("_" if ch in invalid_chars else ch for ch in raw_title).strip()
    return cleaned[:31] or "Hoja"


def excel_column_name(column_index: int) -> str:
    name = ""
    while column_index > 0:
        column_index, remainder = divmod(column_index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def excel_cell_xml(row_index: int, column_index: int, value: object) -> str:
    cell_ref = f"{excel_column_name(column_index)}{row_index}"
    if isinstance(value, bool):
        return f'<c r="{cell_ref}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{cell_ref}"><v>{value}</v></c>'
    text = "" if value is None else str(value)
    escaped = escape(text, {'"': "&quot;"})
    return f'<c r="{cell_ref}" t="inlineStr"><is><t xml:space="preserve">{escaped}</t></is></c>'


def excel_worksheet_xml(rows: List[List[object]]) -> str:
    row_xml: List[str] = []
    for row_index, row in enumerate(rows, start=1):
        cells = "".join(excel_cell_xml(row_index, column_index, value) for column_index, value in enumerate(row, start=1))
        row_xml.append(f'<row r="{row_index}">{cells}</row>')
    sheet_data = "".join(row_xml)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{sheet_data}</sheetData>"
        "</worksheet>"
    )


def excel_workbook_xml(sheet_names: List[str]) -> str:
    sheets_xml = "".join(
        f'<sheet name="{escape(name, {"\"": "&quot;"})}" sheetId="{idx}" r:id="rId{idx}"/>'
        for idx, name in enumerate(sheet_names, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets_xml}</sheets>"
        "</workbook>"
    )


def excel_workbook_rels_xml(sheet_count: int) -> str:
    sheet_rels = "".join(
        f'<Relationship Id="rId{idx}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{idx}.xml"/>'
        for idx in range(1, sheet_count + 1)
    )
    style_rel_id = sheet_count + 1
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{sheet_rels}"
        f'<Relationship Id="rId{style_rel_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        "</Relationships>"
    )


def excel_content_types_xml(sheet_count: int) -> str:
    sheet_overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{idx}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for idx in range(1, sheet_count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        f"{sheet_overrides}"
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )


def write_simple_xlsx(sheets: List[Tuple[str, List[List[object]]]], output_path: Path) -> None:
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )
    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as xlsx_file:
        xlsx_file.writestr("[Content_Types].xml", excel_content_types_xml(len(sheets)))
        xlsx_file.writestr("_rels/.rels", root_rels_xml)
        xlsx_file.writestr("xl/workbook.xml", excel_workbook_xml([sheet_name for sheet_name, _ in sheets]))
        xlsx_file.writestr("xl/_rels/workbook.xml.rels", excel_workbook_rels_xml(len(sheets)))
        xlsx_file.writestr("xl/styles.xml", styles_xml)
        for idx, (_, rows) in enumerate(sheets, start=1):
            xlsx_file.writestr(f"xl/worksheets/sheet{idx}.xml", excel_worksheet_xml(rows))


def write_excel_output(
    nr_policy: str,
    bonds: List[BondSpec],
    bond_summaries: List[BondSummary],
    portfolio_summary: PortfolioSummary,
    pmf_conv: Dict[int, float],
    pmf_brute: Dict[int, float],
    scenario_rows: List[Tuple[Tuple[str, ...], float, float, int]],
    convolution_steps: List[Tuple[str, Dict[int, float]]],
    output_path: Path,
) -> None:
    sheets: List[Tuple[str, List[List[object]]]] = []

    summary_rows: List[List[object]] = [["metric", "value"]]
    summary_rows.extend([[metric, value] for metric, value in portfolio_metric_rows(nr_policy, portfolio_summary)])
    sheets.append(("resumen_portafolio", summary_rows))

    bond_summary_rows_excel: List[List[object]] = [
        ["bond", "rating", "current_price", "expected_horizon_value", "expected_loss", "volatility", "var_999"]
    ]
    bond_summary_rows_excel.extend([list(row) for row in bond_summary_rows(bond_summaries)])
    sheets.append(("resumen_bonos", bond_summary_rows_excel))

    for summary in bond_summaries:
        bond_rows: List[List[object]] = [["state", "probability", "horizon_value", "loss_vs_current"]]
        bond_rows.extend(
            [[row.state, row.probability, row.horizon_value, row.loss_vs_current] for row in summary.distribution]
        )
        sheets.append((excel_sheet_title(summary.bond.name), bond_rows))

    conv_rows: List[List[object]] = [["portfolio_value", "probability", "cumulative_probability"]]
    conv_rows.extend([list(row) for row in portfolio_distribution_rows(pmf_conv)])
    sheets.append(("portafolio_distrib", conv_rows))

    brute_pmf_rows: List[List[object]] = [["portfolio_value", "probability", "cumulative_probability"]]
    brute_pmf_rows.extend([list(row) for row in portfolio_distribution_rows(pmf_brute)])
    sheets.append(("bruteforce_pmf", brute_pmf_rows))

    comparison_rows_excel: List[List[object]] = [
        [
            "portfolio_value",
            "probability_convolution",
            "probability_bruteforce",
            "difference",
            "cumulative_convolution",
            "cumulative_bruteforce",
        ]
    ]
    comparison_rows_excel.extend([list(row) for row in pmf_comparison_rows(pmf_conv, pmf_brute)])
    sheets.append(("comparacion_pmf", comparison_rows_excel))

    brute_rows: List[List[object]] = [[*([bond.name for bond in bonds]), "probability", "portfolio_value", "portfolio_value_cents"]]
    brute_rows.extend([list(states) + [prob, value, value_cents] for states, prob, value, value_cents in scenario_rows])
    sheets.append(("bruteforce_escenarios", brute_rows))

    for step_name, step_pmf in convolution_steps:
        step_rows: List[List[object]] = [["portfolio_value", "probability", "cumulative_probability"]]
        step_rows.extend([list(row) for row in portfolio_distribution_rows(step_pmf)])
        sheets.append((excel_sheet_title(step_name), step_rows))

    tail_rows: List[List[object]] = [["metric", "value"]]
    tail_rows.extend([[metric, value] for metric, value in portfolio_metric_rows(nr_policy, portfolio_summary)])
    sheets.append(("cola_0_1pct", tail_rows))

    write_simple_xlsx(sheets, output_path)


def safe_filename(raw_name: str) -> str:
    cleaned = []
    for ch in raw_name.lower():
        if ch.isalnum():
            cleaned.append(ch)
        elif ch in {" ", "-", "_"}:
            cleaned.append("_")
    return "".join(cleaned).strip("_") or "archivo"


def set_visual_theme() -> None:
    plt.style.use("seaborn-v0_8-whitegrid")
    plt.rcParams.update(
        {
            "axes.facecolor": "#FFFFFF",
            "figure.facecolor": "#FFFFFF",
            "grid.color": "#D9E2EC",
            "grid.linewidth": 0.8,
            "axes.edgecolor": "#BCCCDC",
            "axes.labelcolor": "#243B53",
            "xtick.color": "#243B53",
            "ytick.color": "#243B53",
            "axes.titlecolor": "#102A43",
        }
    )


def format_table_value(value: object) -> str:
    if isinstance(value, float):
        return f"{value:,.6f}"
    if isinstance(value, bool):
        return "Sí" if value else "No"
    return str(value)


def metric_display_name(metric_name: str) -> str:
    if metric_name == "nr_policy":
        return "Política NR"
    if metric_name == "var_999":
        return "VaR 99.9%"
    if metric_name == "es_999":
        return "ES 99.9%"
    if metric_name == "lower_tail_value_001":
        return "Valor cuantil 0.1%"
    return metric_name.replace("_", " ").title()


def render_table_figure(
    title: str,
    columns: List[str],
    rows: List[List[object]],
    subtitle: str | None = None,
    figsize: Tuple[float, float] | None = None,
) -> plt.Figure:
    formatted_rows = [[format_table_value(value) for value in row] for row in rows]
    n_rows = max(len(formatted_rows), 1)
    n_cols = max(len(columns), 1)
    if figsize is None:
        fig_width = max(10.0, min(16.0, 1.8 * n_cols))
        fig_height = max(3.2, min(14.0, 0.42 * (n_rows + 3)))
    else:
        fig_width, fig_height = figsize

    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.axis("off")
    ax.set_title(title, loc="left", fontsize=18, fontweight="bold", pad=12)
    if subtitle:
        ax.text(0.0, 1.02, subtitle, transform=ax.transAxes, fontsize=11, color="#52606D", va="bottom")

    table = ax.table(
        cellText=formatted_rows or [["" for _ in columns]],
        colLabels=columns,
        cellLoc="center",
        colLoc="center",
        bbox=[0.0, 0.0, 1.0, 0.90],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(9.5)
    table.scale(1.0, 1.35)
    try:
        table.auto_set_column_width(col=list(range(len(columns))))
    except Exception:
        pass

    for (row_idx, _), cell in table.get_celld().items():
        cell.set_edgecolor("#D0D7DE")
        if row_idx == 0:
            cell.set_facecolor("#E8F1F8")
            cell.set_text_props(weight="bold", color="#102A43")
        else:
            cell.set_facecolor("#FFFFFF" if row_idx % 2 else "#F8FAFC")

    return fig


def save_figure_bundle(fig: plt.Figure, image_path: Path, pdf_pages: PdfPages) -> None:
    image_path.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(image_path, dpi=220, bbox_inches="tight")
    pdf_pages.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def portfolio_tail_rows(pmf_conv: Dict[int, float], limit: int = 25) -> List[List[object]]:
    return [list(row) for row in portfolio_distribution_rows(pmf_conv)[:limit]]


def comparison_tail_rows(pmf_conv: Dict[int, float], pmf_brute: Dict[int, float], limit: int = 25) -> List[List[object]]:
    return [list(row) for row in pmf_comparison_rows(pmf_conv, pmf_brute)[:limit]]


def worst_scenario_rows(
    bonds: List[BondSpec],
    scenario_rows: List[Tuple[Tuple[str, ...], float, float, int]],
    current_portfolio: float,
    limit: int = 25,
) -> Tuple[List[str], List[List[object]]]:
    columns = [bond.name for bond in bonds] + ["probability", "portfolio_value", "portfolio_value_cents", "loss_vs_current"]
    worst_rows = []
    for states, prob, value, value_cents in sorted(scenario_rows, key=lambda item: (item[2], item[0]))[:limit]:
        worst_rows.append([*states, prob, value, value_cents, current_portfolio - value])
    return columns, worst_rows


def create_cover_figure(nr_policy: str, portfolio_summary: PortfolioSummary) -> plt.Figure:
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.axis("off")
    ax.text(0.02, 0.90, "CreditMetrics", fontsize=28, fontweight="bold", color="#102A43", transform=ax.transAxes)
    ax.text(
        0.02,
        0.82,
        "Reporte visual de distribuciones, tablas y comparación entre convolución y brute force",
        fontsize=14,
        color="#486581",
        transform=ax.transAxes,
    )

    summary_lines = [
        f"Política NR: {nr_policy}",
        f"Valor actual del portafolio: {portfolio_summary.current_value:,.6f}",
        f"Valor esperado a 1 año: {portfolio_summary.expected_horizon_value:,.6f}",
        f"Pérdida esperada: {portfolio_summary.expected_loss:,.6f}",
        f"Volatilidad: {portfolio_summary.volatility:,.6f}",
        f"VaR 99.9%: {portfolio_summary.var_999:,.6f}",
        f"ES 99.9%: {portfolio_summary.es_999:,.6f}",
        f"Cuantil inferior 0.1%: {portfolio_summary.lower_tail_value_001:,.6f}",
        f"Escenarios brute force: {portfolio_summary.support_size_bruteforce:,}",
        f"Soporte por convolución: {portfolio_summary.support_size_cents:,}",
        f"Convolución = brute force: {'Sí' if portfolio_summary.convolution_matches_bruteforce_cents else 'No'}",
        f"Escenario umbral: {' | '.join(portfolio_summary.quantile_state_example)}",
    ]
    ax.text(
        0.03,
        0.70,
        "\n".join(summary_lines),
        fontsize=13,
        va="top",
        linespacing=1.55,
        color="#243B53",
        transform=ax.transAxes,
    )
    ax.text(
        0.03,
        0.10,
        "Estilo visual: seaborn whitegrid. Este PDF resume las salidas más útiles para entregar el trabajo.",
        fontsize=11,
        color="#52606D",
        transform=ax.transAxes,
    )
    return fig


def create_bond_distribution_figure(summary: BondSummary) -> plt.Figure:
    states = [row.state for row in summary.distribution]
    probabilities = [row.probability for row in summary.distribution]
    horizon_values = [row.horizon_value for row in summary.distribution]
    losses = [row.loss_vs_current for row in summary.distribution]
    colors = ["#C0392B" if state == "D" else "#4C78A8" for state in states]

    fig, axes = plt.subplots(2, 1, figsize=(14, 9), sharex=True, gridspec_kw={"height_ratios": [1.0, 1.15]})
    fig.suptitle(
        f"{summary.bond.name} | rating inicial {summary.bond.rating}",
        fontsize=18,
        fontweight="bold",
        y=0.98,
    )

    axes[0].bar(states, probabilities, color=colors, edgecolor="#1F2933", linewidth=0.3)
    axes[0].yaxis.set_major_formatter(PercentFormatter(1.0))
    axes[0].set_ylabel("Probabilidad")
    axes[0].set_title("Distribución de probabilidades por estado migrado", loc="left", fontsize=12)

    axes[1].bar(states, horizon_values, color=colors, alpha=0.88, edgecolor="#1F2933", linewidth=0.3)
    axes[1].axhline(summary.current_price, color="#C0392B", linestyle="--", linewidth=2, label="Precio actual")
    axes[1].axhline(
        summary.expected_horizon_value,
        color="#117A65",
        linestyle="-.",
        linewidth=2,
        label="Valor esperado a 1 año",
    )
    axes[1].set_ylabel("Valor al horizonte")
    axes[1].set_title("Valor al horizonte por estado", loc="left", fontsize=12)

    loss_axis = axes[1].twinx()
    loss_axis.plot(states, losses, color="#D35400", marker="o", linewidth=1.6, label="Pérdida vs actual")
    loss_axis.set_ylabel("Pérdida vs actual")

    handles_left, labels_left = axes[1].get_legend_handles_labels()
    handles_right, labels_right = loss_axis.get_legend_handles_labels()
    axes[1].legend(handles_left + handles_right, labels_left + labels_right, loc="upper right")
    axes[1].tick_params(axis="x", rotation=45)
    return fig


def create_portfolio_distribution_figure(
    pmf_conv: Dict[int, float],
    pmf_brute: Dict[int, float],
    portfolio_summary: PortfolioSummary,
) -> plt.Figure:
    conv_rows = portfolio_distribution_rows(pmf_conv)
    brute_rows = portfolio_distribution_rows(pmf_brute)

    values_conv = [row[0] for row in conv_rows]
    prob_conv = [row[1] for row in conv_rows]
    cdf_conv = [row[2] for row in conv_rows]

    values_brute = [row[0] for row in brute_rows]
    prob_brute = [row[1] for row in brute_rows]
    cdf_brute = [row[2] for row in brute_rows]

    fig, axes = plt.subplots(2, 1, figsize=(14, 10), sharex=True)
    fig.suptitle("Portafolio: convolución vs brute force", fontsize=18, fontweight="bold", y=0.98)

    axes[0].plot(values_conv, prob_conv, linewidth=2.4, color="#2A9D8F", label="Convolución")
    axes[0].plot(values_brute, prob_brute, linewidth=1.8, linestyle="--", color="#E76F51", label="Brute force")
    axes[0].axvline(
        portfolio_summary.lower_tail_value_001,
        color="#6C5CE7",
        linestyle=":",
        linewidth=2,
        label="Cuantil 0.1%",
    )
    axes[0].set_ylabel("Probabilidad")
    axes[0].set_title("PMF del valor del portafolio", loc="left", fontsize=12)
    axes[0].legend(loc="upper right")

    axes[1].plot(values_conv, cdf_conv, linewidth=2.4, color="#2A9D8F", label="Convolución")
    axes[1].plot(values_brute, cdf_brute, linewidth=1.8, linestyle="--", color="#E76F51", label="Brute force")
    axes[1].axhline(0.001, color="#6C5CE7", linestyle=":", linewidth=2, label="0.1% acumulado")
    axes[1].axvline(portfolio_summary.lower_tail_value_001, color="#6C5CE7", linestyle=":", linewidth=2)
    axes[1].set_xlabel("Valor del portafolio")
    axes[1].set_ylabel("Probabilidad acumulada")
    axes[1].set_title("CDF del valor del portafolio", loc="left", fontsize=12)
    axes[1].legend(loc="lower right")
    return fig


def create_convolution_steps_figure(convolution_steps: List[Tuple[str, Dict[int, float]]]) -> plt.Figure:
    fig, axes = plt.subplots(len(convolution_steps), 1, figsize=(14, 4.2 * len(convolution_steps)), sharex=False)
    if len(convolution_steps) == 1:
        axes = [axes]

    fig.suptitle("Pasos intermedios de la convolución", fontsize=18, fontweight="bold", y=0.995)
    base_palette = list(plt.get_cmap("tab10").colors)
    palette = [base_palette[idx % len(base_palette)] for idx in range(len(convolution_steps))]
    for axis, (color, (step_name, step_pmf)) in zip(axes, zip(palette, convolution_steps)):
        rows = portfolio_distribution_rows(step_pmf)
        values = [row[0] for row in rows]
        probabilities = [row[1] for row in rows]
        axis.plot(values, probabilities, linewidth=2.2, color=color)
        axis.fill_between(values, probabilities, alpha=0.18, color=color)
        axis.set_ylabel("Probabilidad")
        axis.set_title(
            f"{step_name.replace('_', ' ')} | soporte = {len(step_pmf):,}",
            loc="left",
            fontsize=12,
        )
    axes[-1].set_xlabel("Valor acumulado")
    return fig


def write_visual_outputs(result: Dict[str, object], output_dir: str | Path) -> None:
    output_path = Path(output_dir)
    figures_path = output_path / "figuras"
    figures_path.mkdir(parents=True, exist_ok=True)

    set_visual_theme()

    bond_summaries: List[BondSummary] = result["bond_summaries"]  # type: ignore[assignment]
    portfolio_summary: PortfolioSummary = result["portfolio_summary"]  # type: ignore[assignment]
    pmf_conv: Dict[int, float] = result["portfolio_pmf_cents"]  # type: ignore[assignment]
    pmf_brute: Dict[int, float] = result["portfolio_pmf_bruteforce_cents"]  # type: ignore[assignment]
    scenario_rows: List[Tuple[Tuple[str, ...], float, float, int]] = result["bruteforce_scenarios"]  # type: ignore[assignment]
    convolution_steps: List[Tuple[str, Dict[int, float]]] = result["convolution_steps"]  # type: ignore[assignment]
    bonds: List[BondSpec] = result["bonds"]  # type: ignore[assignment]
    nr_policy: str = result["nr_policy"]  # type: ignore[assignment]

    pdf_path = output_path / "reporte_creditmetrics.pdf"
    with PdfPages(pdf_path) as pdf_pages:
        metadata = pdf_pages.infodict()
        metadata["Title"] = "CreditMetrics - Reporte visual"
        metadata["Subject"] = "Distribuciones, tablas y comparación convolución vs brute force"

        save_figure_bundle(create_cover_figure(nr_policy, portfolio_summary), figures_path / "00_portada_resumen.png", pdf_pages)

        portfolio_rows = [[metric_display_name(metric), value] for metric, value in portfolio_metric_rows(nr_policy, portfolio_summary)]
        save_figure_bundle(
            render_table_figure(
                "Resumen del portafolio",
                ["Métrica", "Valor"],
                portfolio_rows,
                subtitle="Resumen ejecutivo listo para insertar como imagen en el reporte.",
                figsize=(12, 5.5),
            ),
            figures_path / "01_tabla_resumen_portafolio.png",
            pdf_pages,
        )

        bond_rows = [list(row) for row in bond_summary_rows(bond_summaries)]
        save_figure_bundle(
            render_table_figure(
                "Resumen por bono",
                ["Bono", "Rating", "Precio actual", "Valor esperado", "Pérdida esperada", "Volatilidad", "VaR 99.9%"],
                bond_rows,
                subtitle="Cada fila resume el comportamiento del bono bajo la matriz de transición.",
                figsize=(15, 4.8),
            ),
            figures_path / "02_tabla_resumen_bonos.png",
            pdf_pages,
        )

        for idx, summary in enumerate(bond_summaries, start=3):
            save_figure_bundle(
                create_bond_distribution_figure(summary),
                figures_path / f"{idx:02d}_grafica_{safe_filename(summary.bond.name)}.png",
                pdf_pages,
            )

        save_figure_bundle(
            create_portfolio_distribution_figure(pmf_conv, pmf_brute, portfolio_summary),
            figures_path / "06_grafica_portafolio_distribucion.png",
            pdf_pages,
        )

        save_figure_bundle(
            create_convolution_steps_figure(convolution_steps),
            figures_path / "07_grafica_pasos_convolucion.png",
            pdf_pages,
        )

        save_figure_bundle(
            render_table_figure(
                "Cola inferior del portafolio",
                ["Valor del portafolio", "Probabilidad", "Prob. acumulada"],
                portfolio_tail_rows(pmf_conv),
                subtitle="Primeros 25 valores del soporte ordenados de peor a mejor.",
                figsize=(12, 9.0),
            ),
            figures_path / "08_tabla_cola_portafolio.png",
            pdf_pages,
        )

        worst_columns, worst_rows = worst_scenario_rows(bonds, scenario_rows, portfolio_summary.current_value)
        save_figure_bundle(
            render_table_figure(
                "Peores escenarios brute force",
                worst_columns,
                worst_rows,
                subtitle="Se muestran los 25 escenarios con menor valor del portafolio.",
                figsize=(16, 9.4),
            ),
            figures_path / "09_tabla_peores_escenarios_bruteforce.png",
            pdf_pages,
        )

        save_figure_bundle(
            render_table_figure(
                "Comparación en la cola: convolución vs brute force",
                [
                    "Valor del portafolio",
                    "Prob. convolución",
                    "Prob. brute force",
                    "Diferencia",
                    "Acum. convolución",
                    "Acum. brute force",
                ],
                comparison_tail_rows(pmf_conv, pmf_brute),
                subtitle="Primeras 25 filas del soporte para verificar visualmente que ambas PMF coinciden.",
                figsize=(16, 9.0),
            ),
            figures_path / "10_tabla_comparacion_convolucion_vs_bruteforce.png",
            pdf_pages,
        )


def run_analysis(
    bonds: Iterable[BondSpec] = EXERCISE_BONDS,
    nr_policy: str = "renormalize",
    output_dir: str | Path | None = None,
) -> Dict[str, object]:
    bonds = list(bonds)
    bond_summaries = [summarize_single_bond(bond, nr_policy=nr_policy) for bond in bonds]
    current_portfolio = sum(item.current_price for item in bond_summaries)

    convolution_steps = portfolio_convolution_steps(bonds, nr_policy=nr_policy)
    pmf_conv = convolution_steps[-1][1]
    pmf_brute, scenario_rows = portfolio_distribution_bruteforce(bonds, nr_policy=nr_policy)
    convolution_matches = pmf_close(pmf_conv, pmf_brute)

    horizon_expected = sum((value_cents / 100.0) * prob for value_cents, prob in pmf_conv.items())
    horizon_variance = sum((((value_cents / 100.0) - horizon_expected) ** 2) * prob for value_cents, prob in pmf_conv.items())

    loss_pmf: Dict[float, float] = defaultdict(float)
    for value_cents, prob in pmf_conv.items():
        loss = round(current_portfolio - (value_cents / 100.0), 2)
        loss_pmf[loss] += prob

    sorted_values = sorted(pmf_conv.items(), key=lambda x: x[0])
    cumulative = 0.0
    lower_tail_value = sorted_values[-1][0] / 100.0
    for value_cents, prob in sorted_values:
        cumulative += prob
        if cumulative >= 0.001:
            lower_tail_value = value_cents / 100.0
            break

    losses_sorted = sorted(loss_pmf.items(), key=lambda x: x[0])
    var_999 = discrete_quantile(losses_sorted, 0.999)
    es_999 = expected_shortfall_from_loss_pmf(loss_pmf, 0.999)

    sorted_scenarios = sorted(scenario_rows, key=lambda x: x[2])
    cumulative_scen = 0.0
    quantile_state_example = sorted_scenarios[-1][0]
    for states, prob, value, value_cents in sorted_scenarios:
        cumulative_scen += prob
        if cumulative_scen >= 0.001:
            quantile_state_example = states
            break

    portfolio_summary = PortfolioSummary(
        current_value=current_portfolio,
        expected_horizon_value=horizon_expected,
        expected_loss=current_portfolio - horizon_expected,
        volatility=math.sqrt(horizon_variance),
        var_999=var_999,
        es_999=es_999,
        lower_tail_value_001=lower_tail_value,
        support_size_cents=len(pmf_conv),
        support_size_bruteforce=len(scenario_rows),
        convolution_matches_bruteforce_cents=convolution_matches,
        quantile_state_example=quantile_state_example,
    )

    result = {
        "nr_policy": nr_policy,
        "bonds": bonds,
        "bond_summaries": bond_summaries,
        "portfolio_summary": portfolio_summary,
        "portfolio_pmf_cents": pmf_conv,
        "portfolio_pmf_bruteforce_cents": pmf_brute,
        "bruteforce_scenarios": scenario_rows,
        "convolution_steps": convolution_steps,
    }

    if output_dir is not None:
        write_outputs(result, output_dir)

    return result


def write_outputs(result: Dict[str, object], output_dir: str | Path) -> None:
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    bond_summaries: List[BondSummary] = result["bond_summaries"]  # type: ignore[assignment]
    portfolio_summary: PortfolioSummary = result["portfolio_summary"]  # type: ignore[assignment]
    pmf_conv: Dict[int, float] = result["portfolio_pmf_cents"]  # type: ignore[assignment]
    pmf_brute: Dict[int, float] = result["portfolio_pmf_bruteforce_cents"]  # type: ignore[assignment]
    scenario_rows: List[Tuple[Tuple[str, ...], float, float, int]] = result["bruteforce_scenarios"]  # type: ignore[assignment]
    convolution_steps: List[Tuple[str, Dict[int, float]]] = result["convolution_steps"]  # type: ignore[assignment]
    bonds: List[BondSpec] = result["bonds"]  # type: ignore[assignment]

    for summary in bond_summaries:
        csv_path = output_path / f"{summary.bond.name.lower()}_distribucion.csv"
        with csv_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["state", "probability", "horizon_value", "loss_vs_current"])
            for row in summary.distribution:
                writer.writerow([
                    row.state,
                    f"{row.probability:.10f}",
                    f"{row.horizon_value:.6f}",
                    f"{row.loss_vs_current:.6f}",
                ])

    pmf_csv = output_path / "portafolio_distribucion.csv"
    with pmf_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["portfolio_value", "probability", "cumulative_probability"])
        for value, prob, cumulative in portfolio_distribution_rows(pmf_conv):
            writer.writerow([f"{value:.2f}", f"{prob:.12f}", f"{cumulative:.12f}"])

    brute_pmf_csv = output_path / "portafolio_distribucion_bruteforce.csv"
    with brute_pmf_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["portfolio_value", "probability", "cumulative_probability"])
        for value, prob, cumulative in portfolio_distribution_rows(pmf_brute):
            writer.writerow([f"{value:.2f}", f"{prob:.12f}", f"{cumulative:.12f}"])

    comparison_csv = output_path / "comparacion_convolucion_vs_bruteforce.csv"
    with comparison_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "portfolio_value",
                "probability_convolution",
                "probability_bruteforce",
                "difference",
                "cumulative_convolution",
                "cumulative_bruteforce",
            ]
        )
        for row in pmf_comparison_rows(pmf_conv, pmf_brute):
            writer.writerow([f"{row[0]:.2f}", f"{row[1]:.12f}", f"{row[2]:.12f}", f"{row[3]:.12e}", f"{row[4]:.12f}", f"{row[5]:.12f}"])

    brute_scenarios_csv = output_path / "bruteforce_escenarios.csv"
    with brute_scenarios_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([bond.name for bond in bonds] + ["probability", "portfolio_value", "portfolio_value_cents"])
        for states, prob, value, value_cents in scenario_rows:
            writer.writerow([*states, f"{prob:.12f}", f"{value:.6f}", value_cents])

    for step_name, step_pmf in convolution_steps:
        step_csv = output_path / f"{step_name}.csv"
        with step_csv.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["portfolio_value", "probability", "cumulative_probability"])
            for value, prob, cumulative in portfolio_distribution_rows(step_pmf):
                writer.writerow([f"{value:.2f}", f"{prob:.12f}", f"{cumulative:.12f}"])

    tail_csv = output_path / "cola_0_1pct.csv"
    with tail_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["metric", "value"])
        for metric, value in portfolio_metric_rows(result["nr_policy"], portfolio_summary):
            if isinstance(value, float):
                writer.writerow([metric, f"{value:.6f}"])
            else:
                writer.writerow([metric, value])

    json_path = output_path / "summary.json"
    payload = {
        "nr_policy": result["nr_policy"],
        "bond_summaries": [
            {
                "bond": asdict(summary.bond),
                "current_price": summary.current_price,
                "expected_horizon_value": summary.expected_horizon_value,
                "expected_loss": summary.expected_loss,
                "volatility": summary.volatility,
                "var_999": summary.var_999,
            }
            for summary in bond_summaries
        ],
        "portfolio_summary": {
            "current_value": portfolio_summary.current_value,
            "expected_horizon_value": portfolio_summary.expected_horizon_value,
            "expected_loss": portfolio_summary.expected_loss,
            "volatility": portfolio_summary.volatility,
            "var_999": portfolio_summary.var_999,
            "es_999": portfolio_summary.es_999,
            "lower_tail_value_001": portfolio_summary.lower_tail_value_001,
            "support_size_cents": portfolio_summary.support_size_cents,
            "support_size_bruteforce": portfolio_summary.support_size_bruteforce,
            "convolution_matches_bruteforce_cents": portfolio_summary.convolution_matches_bruteforce_cents,
            "quantile_state_example": list(portfolio_summary.quantile_state_example),
        },
    }
    json_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

    xlsx_path = output_path / "resultados_creditmetrics.xlsx"
    write_excel_output(
        result["nr_policy"],
        bonds,
        bond_summaries,
        portfolio_summary,
        pmf_conv,
        pmf_brute,
        scenario_rows,
        convolution_steps,
        xlsx_path,
    )
    write_visual_outputs(result, output_path)
