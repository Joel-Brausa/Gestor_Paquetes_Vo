import sys
from pathlib import Path

# Add Voest_Paquets directory to path so modules can be imported
# conftest.py is in .claude/worktrees/jolly-dhawan-bf571a/tests/
# We need to go up 5 levels: tests -> jolly-dhawan-bf571a -> worktrees -> .claude -> Voest_Paquets
voest_paquets_dir = Path(__file__).parent.parent.parent.parent.parent
sys.path.insert(0, str(voest_paquets_dir))
