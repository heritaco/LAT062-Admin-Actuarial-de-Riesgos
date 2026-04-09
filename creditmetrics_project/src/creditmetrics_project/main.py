from __future__ import annotations

import argparse
from pathlib import Path

from .model import run_analysis


def main() -> None:
    parser = argparse.ArgumentParser(description="CreditMetrics exacto por convolución para 3 bonos independientes.")
    parser.add_argument(
        "--nr-policy",
        choices=["renormalize", "keep_current"],
        default="renormalize",
        help="Tratamiento de la masa NR eliminada de la tabla de transición.",
    )
    parser.add_argument(
        "--output-dir",
        default="outputs",
        help="Directorio donde se escribirán CSV/JSON de salida.",
    )
    args = parser.parse_args()

    result = run_analysis(nr_policy=args.nr_policy, output_dir=Path(args.output_dir))
    portfolio = result["portfolio_summary"]
    bonds = result["bond_summaries"]

    print("=" * 72)
    print("CreditMetrics — Portafolio de 3 bonos independientes")
    print("=" * 72)
    print(f"Política NR                : {result['nr_policy']}")
    print(f"Valor actual portafolio    : {portfolio.current_value:,.6f}")
    print(f"Valor esperado a 1 año     : {portfolio.expected_horizon_value:,.6f}")
    print(f"Pérdida esperada           : {portfolio.expected_loss:,.6f}")
    print(f"Volatilidad                : {portfolio.volatility:,.6f}")
    print(f"VaR 99.9%                  : {portfolio.var_999:,.6f}")
    print(f"ES 99.9%                   : {portfolio.es_999:,.6f}")
    print(f"Valor en cuantil 0.1%      : {portfolio.lower_tail_value_001:,.6f}")
    print(f"Escenario umbral ejemplo   : {' | '.join(portfolio.quantile_state_example)}")
    print(f"Soporte por convolución    : {portfolio.support_size_cents:,}")
    print(f"Escenarios brute force     : {portfolio.support_size_bruteforce:,}")
    print(f"Convolución = brute force? : {portfolio.convolution_matches_bruteforce_cents}")
    print()
    print("Resumen por bono")
    print("-" * 72)
    for item in bonds:
        print(
            f"{item.bond.name:18s} rating={item.bond.rating:4s}  P0={item.current_price:10.6f}  "
            f"E[V1]={item.expected_horizon_value:10.6f}  VaR99.9={item.var_999:10.6f}"
        )
    print("=" * 72)


if __name__ == "__main__":
    main()
