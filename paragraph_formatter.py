"""
Module for handling paragraph formatting based on predefined hierarchical levels.
"""

class ParagraphFormatter:
    """
    Class to handle automatic paragraph formatting based on predefined levels:
    - Level 0: 壹, 貳, 參, 肆, ...
    - Level 1: 一, 二, 三, 四, 五, ...
    - Level 2: (一), (二), (三), (四), (五), ...
    - Level 3: 1., 2., 3., 4., 5., ...
    - Level 4: (1), (2), (3), (4), (5), ...
    """
    
    def __init__(self):
        # Define the characters for each level
        self.level_0_chars = ["壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖", "拾"]
        self.level_1_chars = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
        self.level_2_format = "({0})"  # Format for level 2: (一), (二), etc.
        self.level_3_format = "{0}."   # Format for level 3: 1., 2., etc.
        self.level_4_format = "({0})"  # Format for level 4: (1), (2), etc.
        
        # Initialize counters for each level
        self.reset_counters()
        
        # Define indentation for each level
        self.level_indents = ["", "  ", "    ", "      ", "        "]
    
    def reset_counters(self):
        """Reset all level counters to zero"""
        self.level_counters = [0, 0, 0, 0, 0]
    
    def get_next_marker(self, level, increment=True):
        """
        Get the next marker for the specified level
        
        Args:
            level (int): The level for which to get the next marker (0-4)
            increment (bool): Whether to increment the counter for this level
            
        Returns:
            str: The formatted marker for the specified level
        """
        if level < 0 or level > 4:
            raise ValueError("Level must be between 0 and 4")
        
        # Increment the counter for this level if requested
        if increment:
            self.level_counters[level] += 1
            
            # Reset all deeper level counters
            for i in range(level + 1, 5):
                self.level_counters[i] = 0
        
        # Get the index (1-based for display)
        idx = self.level_counters[level] - 1
        if idx < 0:
            idx = 0
            self.level_counters[level] = 1
        
        # Format based on level
        if level == 0:
            # Level 0: 壹, 貳, 參, 肆, ...
            return self.level_0_chars[idx % len(self.level_0_chars)]
        elif level == 1:
            # Level 1: 一, 二, 三, 四, 五, ...
            return self.level_1_chars[idx % len(self.level_1_chars)]
        elif level == 2:
            # Level 2: (一), (二), (三), (四), (五), ...
            return self.level_2_format.format(self.level_1_chars[idx % len(self.level_1_chars)])
        elif level == 3:
            # Level 3: 1., 2., 3., 4., 5., ...
            return self.level_3_format.format(idx + 1)
        else:  # level == 4
            # Level 4: (1), (2), (3), (4), (5), ...
            return self.level_4_format.format(idx + 1)
    
    def get_current_marker(self, level):
        """
        Get the current marker for the specified level without incrementing
        
        Args:
            level (int): The level for which to get the current marker (0-4)
            
        Returns:
            str: The formatted marker for the specified level
        """
        return self.get_next_marker(level, increment=False)
    
    def format_paragraph(self, text, level):
        """
        Format a paragraph with the appropriate marker and indentation for the specified level
        
        Args:
            text (str): The paragraph text to format
            level (int): The level for the paragraph (0-4)
            
        Returns:
            str: The formatted paragraph with the appropriate marker and indentation
        """
        marker = self.get_next_marker(level)
        indent = self.level_indents[level]
        return f"{indent}{marker} {text}"
    
    def get_indentation(self, level):
        """
        Get the indentation for the specified level
        
        Args:
            level (int): The level for which to get the indentation (0-4)
            
        Returns:
            str: The indentation string for the specified level
        """
        if level < 0 or level > 4:
            return ""
        return self.level_indents[level]
    
    def detect_level(self, line):
        """
        Attempt to detect the level of a line based on its format
        
        Args:
            line (str): The line to analyze
            
        Returns:
            int: The detected level (0-4) or -1 if no level format is detected
            str: The content after the marker, or the original line if no marker is detected
        """
        line = line.strip()
        content = line
        level = -1
        
        # Check for level 0: 壹, 貳, 參, 肆, ...
        if any(line.startswith(char) for char in self.level_0_chars):
            level = 0
            # Find the first space after the marker
            space_idx = line.find(' ')
            if space_idx > 0:
                content = line[space_idx+1:]
            
        # Check for level 1: 一, 二, 三, 四, 五, ...
        elif any(line.startswith(char) for char in self.level_1_chars):
            level = 1
            # Find the first space after the marker
            space_idx = line.find(' ')
            if space_idx > 0:
                content = line[space_idx+1:]
            
        # Check for level 2: (一), (二), (三), (四), (五), ...
        elif any(line.startswith(f"({char})") for char in self.level_1_chars):
            level = 2
            # Find the first space after the marker
            space_idx = line.find(' ')
            if space_idx > 0:
                content = line[space_idx+1:]
            
        # Check for level 3: 1., 2., 3., 4., 5., ...
        elif line and line[0].isdigit() and line.find('.') > 0:
            level = 3
            # Find the first space after the marker
            space_idx = line.find(' ')
            if space_idx > 0:
                content = line[space_idx+1:]
            
        # Check for level 4: (1), (2), (3), (4), (5), ...
        elif line and line.startswith('(') and line.find(')') > 0:
            try:
                int(line[1:line.find(')')])
                level = 4
                # Find the first space after the marker
                space_idx = line.find(' ')
                if space_idx > 0:
                    content = line[space_idx+1:]
            except ValueError:
                pass
                
        return level, content
