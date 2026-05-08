"""Classical Go proverbs for display in the sidebar.

PROVERBS is the canonical list of proverbs.
random_proverb() returns one at random, optionally seeded for reproducibility.
"""

import random as _random

PROVERBS: list[str] = [
    # For all situations
    "The enemy's key point is yours.",
    "Play on the point of symmetry.",
    "Beware of going back to patch up.",
    "When in doubt, Tenuki.",
    # Life-and-death
    "There is death in the hane.",
    "Learn the eye-stealing tesuji.",
    "Six die but eight live (2nd line).",
    "Four die but six live (3rd line).",
    "The L group is dead.",
    "The door group is dead.",
    "Strange things happen at the 1-2 point.",
    "Eyes win semeais.",
    # Tactics
    "Respond to attachment with hane.",
    "Hane at the head of two stones.",
    "Crosscut, then extend.",
    "Capture the cutting stones.",
    "The empty triangle is bad shape.",
    "The one-point jump is never wrong.",
    "Strike at the waist of the keima.",
    "Even a moron connects against a peep.",
    "If you have one stone on the third line in atari, add a second stone and sacrifice both.",
    "Do not peep at cutting points.",
    "Use contact moves for defence.",
    "Never ignore a shoulder hit.",
    "Nets are better than ladders.",
    "The bamboo joint may be short of liberties.",
    "Approach from the wider side.",
    "Answer the capping play with a keima.",
    "Answer keima with kosumi.",
    "Block on the wider side.",
    "Play (one point jump) at the centre of three stones.",
    "Capture stones caught in a ladder at the earliest opportunity.",
    "Play double sente early.",
    # Strategy
    "Urgent points before big points.",
    "Play away from thickness.",
    "Don't use thickness to make territory.",
    "Make territory while attacking.",
    "A ponnuki is worth thirty points.",
    "Make a fist before striking.",
    "A rich man should not pick quarrels.",
    "Don't attach when attacking.",
    "Make weak walk along with weak.",
    "Five groups might live, but the sixth will die.",
    "Big dragons never die.",
    "Give your opponent what he wants.",
    "Sacrifice plums for peaches.",
    "Don't throw an egg at a wall.",
    "Don't go fishing while your house is on fire.",
    "High move for influence, low move for territory.",
    "Don't push from behind.",
    "Don't push along the second line.",
    "Riding the tiger, it is difficult to get off.",
    "Make a feint to the east while attacking in the west.",
    # Meta
    "Don't follow proverbs blindly.",
    "Lose your first fifty games as quickly as possible.",
    "You can play Go, but don't let Go play you.",
    "The threat is stronger than its execution.",
    # General wisdom applied to Go
    "Rice eaten in haste chokes.",
    "When you want to test the depth of a stream, don't use both feet.",
    "Distant water won't help to put out a nearby fire.",
    "The fearless suffer one death; the fearful suffer a thousand.",
]


def random_proverb(seed: int | None = None) -> str:
    """Return a random proverb from PROVERBS.

    Args:
        seed: Optional integer seed for reproducible selection.

    Returns:
        A single proverb string chosen at random.
    """
    rng = _random.Random(seed)
    return rng.choice(PROVERBS)
