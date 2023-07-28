from logic.utils.tree_dirs import TreeDir
from logic.bicu.comparer import BicuComparer
from logic.bicu.calculate import BicuCalculate
from logic.consumers.comparer import ConsumersComparer
from logic.consumers.calculate import Calculate

TreeDir().create_tree_dirs()
BicuCalculate().format_data(5)
