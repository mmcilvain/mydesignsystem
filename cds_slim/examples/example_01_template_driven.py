"""
Template-driven example — the recommended path.

Opens assets/comcast_template.pptx, strips the sample slides, builds a deck
using layouts referenced by slug. The output keeps the Comcast embedded
fonts and pixel-perfect layouts.

Run:
    python -m comcast_design_system.examples.example_01_template_driven
"""
import sys
from pathlib import Path

# add repo root to path so the relative imports work when run as a script
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.deck import Deck


def build():
    deck = Deck()  # samples cleared by default

    deck.add('presentation_title_slide',
             title='Q3 Operating Review',
             subhead='Revenue, product, and the path into Q4',
             date='OCTOBER 2026')

    deck.add('section_title',
             title='Top-line growth')

    deck.add('comcast_standard',
             banner='OPERATING REVIEW',
             title='Where we landed in Q3',
             subhead='Top-line, retention, and signups',
             body=('Revenue grew 14% YoY to $2.4B, the best quarterly print since 2024. '
                   'Net adds came in at 380K against guidance of 340K, '
                   'driven by bundles outperforming streaming churn for the first time in 4 quarters.'))

    deck.add('data_3',
             banner='KEY METRICS',
             title='Three numbers from Q3',
             stat_values=['14%', '$2.4B', '380K'],
             stat_captions=['Revenue growth YoY',
                            'Total revenue',
                            'Net adds vs 340K guide'])

    deck.add('key_points_slide',
             banner='WHY IT MATTERS',
             title='Three things to watch into Q4',
             subheaders=['Pricing power',
                         'Net adds vs guide',
                         'Retention curve',
                         'Streaming mix',
                         'Cost discipline'],
             bodies=['Bundles holding share against streaming-only churn.',
                     'First quarter beating guide in 4 quarters.',
                     '7-day retention up 220bps QoQ.',
                     'Streaming attach rate up 4 points to 38%.',
                     'OpEx growth held to 2% on flat headcount.'])

    deck.add('closing')

    out = Path(__file__).resolve().parent.parent / 'output' / 'q3_review.pptx'
    deck.save(out)
    print(f'wrote {out}  ({deck.slide_count} slides)')


if __name__ == '__main__':
    build()
