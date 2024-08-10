from enum import Enum


class Color(Enum):
    RED = 0x00
    BLACK = 0x01


class Node:
    def __init__(
        self,
        key,
        color=Color.RED,
        left=None,
        right=None,
        parent=None,
    ):
        self.key = key
        self.color = color
        self.left = left
        self.right = right
        self.parent = parent


class RBTree:
    def __init__(self):
        self.NIL = Node(key=None, color=Color.BLACK)  # Sentinel NIL node
        self.root = self.NIL

    def insert(self, key):
        new_node = Node(key=key, left=self.NIL, right=self.NIL)
        parent = None
        current = self.root

        while current != self.NIL:
            parent = current

            if new_node.key < current.key:
                current = current.left

            else:
                current = current.right

        new_node.parent = parent

        # Tree is empty
        if parent is None:
            self.root = new_node

        elif new_node.key < parent.key:
            parent.left = new_node

        else:
            parent.right = new_node

        new_node.left = self.NIL
        new_node.right = self.NIL
        new_node.color = Color.RED

        self.fix_insert(new_node)

    def fix_insert(self, node):
        while node.parent and node.parent.color == Color.RED:
            if node.parent == node.parent.parent.left:
                uncle = node.parent.parent.right

                # Case 1: Uncle is red
                if uncle.color == Color.RED:
                    node.parent.color = Color.BLACK
                    uncle.color = Color.BLACK
                    node.parent.parent.color = Color.RED
                    node = node.parent.parent

                else:
                    # Case 2: Node is right child
                    if node == node.parent.right:
                        node = node.parent
                        self.left_rotate(node)

                    # Case 3: Node is left child
                    node.parent.color = Color.BLACK
                    node.parent.parent.color = Color.RED
                    self.right_rotate(node.parent.parent)

            else:
                uncle = node.parent.parent.left

                # Case 1: Uncle is red
                if uncle.color == Color.RED:
                    node.parent.color = Color.BLACK
                    uncle.color = Color.BLACK
                    node.parent.parent.color = Color.RED
                    node = node.parent.parent

                else:
                    # Case 2: Node is left child
                    if node == node.parent.left:
                        node = node.parent
                        self.right_rotate(node)

                    # Case 3: Node is right child
                    node.parent.color = Color.BLACK
                    node.parent.parent.color = Color.RED
                    self.left_rotate(node.parent.parent)

        self.root.color = Color.BLACK

    def left_rotate(self, node):
        right_child = node.right
        node.right = right_child.left

        if right_child.left != self.NIL:
            right_child.left.parent = node

        right_child.parent = node.parent

        # Node is root
        if node.parent is None:
            self.root = right_child

        elif node == node.parent.left:
            node.parent.left = right_child

        else:
            node.parent.right = right_child

        right_child.left = node
        node.parent = right_child

    def right_rotate(self, node):
        left_child = node.left
        node.left = left_child.right

        if left_child.right != self.NIL:
            left_child.right.parent = node

        left_child.parent = node.parent

        if node.parent is None:  # Node is root
            self.root = left_child

        elif node == node.parent.right:
            node.parent.right = left_child

        else:
            node.parent.left = left_child

        left_child.right = node
        node.parent = left_child

    def search(self, key):
        return self._search_tree_helper(self.root, key)

    def _search_tree_helper(self, node, key):
        if node == self.NIL or key == node.key:
            return node

        if key < node.key:
            return self._search_tree_helper(node.left, key)

        return self._search_tree_helper(node.right, key)

    def inorder(self, node):
        if node != self.NIL:
            self.inorder(node.left)
            print(f"{node.key} ({node.color})", end=" ")
            self.inorder(node.right)


# Example usage
if __name__ == "__main__":
    tree = RBTree()
    keys = [20, 15, 25, 10, 5, 1, 30, 40]

    for key in keys:
        tree.insert(key)

    print("Inorder traversal of the tree:")
    tree.inorder(tree.root)

    # Search for a key
    key_to_search = 15
    result = tree.search(key_to_search)
    if result != tree.NIL:
        print(f"\nKey {key_to_search} found in the tree.")
    else:
        print(f"\nKey {key_to_search} not found in the tree.")
