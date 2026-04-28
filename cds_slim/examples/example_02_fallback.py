"""
Fallback example — build a deck without the .pptx template.

Use this when the .pptx isn't available (CI runners, restricted envs).
Output is on-palette but won't have Comcast New Vision (uses Aptos fallback)
and won't have peacock graphics. Good enough for internal automation.

Run:
    python -m comcast_design_system.examples.example_02_fallback
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import brand
from src.fallback import BlankDeck


def build():
    d = BlankDeck()

    d.add_title_slide(title='Q3 Operating Review',
                      subhead='Revenue, product, and the path into Q4',
                      date='OCTOBER 2026')

    # default accent is yellow — pass any brand color from src.brand.COLORS
    d.add_section_divider(eyebrow='SECTION 1',
                          title='Top-line growth',
                          accent_hex=brand.COLORS.blue)

    d.add_standard_slide(banner='OPERATING REVIEW',
                         title='Where we landed in Q3',
                         subhead='Top-line, retention, and signups',
                         body=('Revenue grew 14% YoY to $2.4B. Net adds came in at 380K '
                               'against guidance of 340K, with bundles outperforming '
                               'streaming churn for the first time in 4 quarters.'))

    d.add_data_slide(banner='KEY METRICS',
                     title='Three numbers from Q3',
                     stats=[('14%',   'Revenue growth YoY'),
                            ('$2.4B', 'Total revenue'),
                            ('380K',  'Net adds vs 340K guide')])

    d.add_key_points_slide(banner='WHY IT MATTERS',
                           title='Three things to watch',
                           points=[('Pricing power', 'Bundles holding share against streaming churn.'),
                                   ('Net adds vs guide', 'First beat in 4 quarters.'),
                                   ('Retention curve', '7-day retention up 220bps QoQ.')])

    d.add_closing_slide()

    out = Path(__file__).resolve().parent.parent / 'output' / 'q3_review_fallback.pptx'
    d.save(out)
    print(f'wrote {out}  ({d.slide_count} slides)')


if __name__ == '__main__':
    build()
