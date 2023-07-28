from logic.utils.tree_dirs import TreeDir
from logic.bicu.comparer import BicuComparer
from logic.consumers.comparer import ConsumersComparer
from logic.consumers.calculate import Calculate

TreeDir().create_tree_dirs()
Calculate().format_data(5, "Коммерческих")
