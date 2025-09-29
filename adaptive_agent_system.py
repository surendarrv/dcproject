"""
Adaptive Agent System - Learns from User Feedback and Auto-corrects Output

This system demonstrates how to deploy a program as an agent that:
1. Generates initial outputs
2. Captures user corrections through UI
3. Learns from those corrections
4. Auto-corrects future outputs based on learned patterns
"""

import json
import sqlite3
import numpy as np
from datetime import datetime
from typing import Dict, List, Any, Tuple
from dataclasses import dataclass
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import pickle
import os

@dataclass
class UserFeedback:
    """Represents user feedback on agent output"""
    original_output: str
    corrected_output: str
    user_id: str
    timestamp: datetime
    context: Dict[str, Any]
    feedback_type: str  # 'correction', 'approval', 'rejection'

@dataclass
class LearningPattern:
    """Represents a learned pattern from user feedback"""
    pattern_id: str
    input_features: List[str]
    correction_mapping: Dict[str, str]
    confidence_score: float
    usage_count: int
    last_updated: datetime

class FeedbackDatabase:
    """Manages storage and retrieval of user feedback"""
    
    def __init__(self, db_path: str = "feedback.db"):
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        """Initialize the feedback database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Create feedback table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS feedback (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                original_output TEXT,
                corrected_output TEXT,
                user_id TEXT,
                timestamp TEXT,
                context TEXT,
                feedback_type TEXT
            )
        ''')
        
        # Create learning patterns table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS learning_patterns (
                pattern_id TEXT PRIMARY KEY,
                input_features TEXT,
                correction_mapping TEXT,
                confidence_score REAL,
                usage_count INTEGER,
                last_updated TEXT
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def store_feedback(self, feedback: UserFeedback):
        """Store user feedback in database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO feedback (original_output, corrected_output, user_id, 
                                timestamp, context, feedback_type)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            feedback.original_output,
            feedback.corrected_output,
            feedback.user_id,
            feedback.timestamp.isoformat(),
            json.dumps(feedback.context),
            feedback.feedback_type
        ))
        
        conn.commit()
        conn.close()
    
    def get_feedback_history(self, user_id: str = None, limit: int = 100) -> List[UserFeedback]:
        """Retrieve feedback history"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        query = "SELECT * FROM feedback"
        params = []
        
        if user_id:
            query += " WHERE user_id = ?"
            params.append(user_id)
        
        query += " ORDER BY timestamp DESC LIMIT ?"
        params.append(limit)
        
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        
        feedback_list = []
        for row in rows:
            feedback_list.append(UserFeedback(
                original_output=row[1],
                corrected_output=row[2],
                user_id=row[3],
                timestamp=datetime.fromisoformat(row[4]),
                context=json.loads(row[5]),
                feedback_type=row[6]
            ))
        
        return feedback_list

class LearningEngine:
    """Core learning engine that processes feedback and generates patterns"""
    
    def __init__(self, db: FeedbackDatabase):
        self.db = db
        self.vectorizer = TfidfVectorizer(max_features=1000, stop_words='english')
        self.patterns: Dict[str, LearningPattern] = {}
        self.load_existing_patterns()
    
    def load_existing_patterns(self):
        """Load existing learning patterns from database"""
        conn = sqlite3.connect(self.db.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM learning_patterns")
        rows = cursor.fetchall()
        conn.close()
        
        for row in rows:
            pattern = LearningPattern(
                pattern_id=row[0],
                input_features=json.loads(row[1]),
                correction_mapping=json.loads(row[2]),
                confidence_score=row[3],
                usage_count=row[4],
                last_updated=datetime.fromisoformat(row[5])
            )
            self.patterns[pattern.pattern_id] = pattern
    
    def process_feedback(self, feedback: UserFeedback):
        """Process new feedback and update learning patterns"""
        # Extract features from the original output
        features = self.extract_features(feedback.original_output)
        
        # Find similar patterns or create new one
        pattern_id = self.find_or_create_pattern(features, feedback)
        
        # Update the pattern with new correction
        self.update_pattern(pattern_id, feedback)
        
        # Save updated patterns to database
        self.save_patterns()
    
    def extract_features(self, text: str) -> List[str]:
        """Extract features from text for pattern matching"""
        # Simple feature extraction - can be enhanced with more sophisticated NLP
        words = text.lower().split()
        return [word for word in words if len(word) > 2]
    
    def find_or_create_pattern(self, features: List[str], feedback: UserFeedback) -> str:
        """Find existing similar pattern or create new one"""
        best_similarity = 0
        best_pattern_id = None
        
        for pattern_id, pattern in self.patterns.items():
            similarity = self.calculate_similarity(features, pattern.input_features)
            if similarity > best_similarity and similarity > 0.7:  # Threshold for similarity
                best_similarity = similarity
                best_pattern_id = pattern_id
        
        if best_pattern_id:
            return best_pattern_id
        else:
            # Create new pattern
            new_pattern_id = f"pattern_{len(self.patterns) + 1}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            return new_pattern_id
    
    def calculate_similarity(self, features1: List[str], features2: List[str]) -> float:
        """Calculate similarity between two feature sets"""
        if not features1 or not features2:
            return 0.0
        
        # Convert to TF-IDF vectors and calculate cosine similarity
        try:
            combined_text = [' '.join(features1), ' '.join(features2)]
            tfidf_matrix = self.vectorizer.fit_transform(combined_text)
            similarity = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0]
            return similarity
        except:
            # Fallback to simple word overlap
            set1, set2 = set(features1), set(features2)
            intersection = len(set1.intersection(set2))
            union = len(set1.union(set2))
            return intersection / union if union > 0 else 0.0
    
    def update_pattern(self, pattern_id: str, feedback: UserFeedback):
        """Update pattern with new feedback"""
        if pattern_id not in self.patterns:
            # Create new pattern
            self.patterns[pattern_id] = LearningPattern(
                pattern_id=pattern_id,
                input_features=self.extract_features(feedback.original_output),
                correction_mapping={},
                confidence_score=0.5,
                usage_count=0,
                last_updated=datetime.now()
            )
        
        pattern = self.patterns[pattern_id]
        
        # Update correction mapping
        if feedback.original_output not in pattern.correction_mapping:
            pattern.correction_mapping[feedback.original_output] = feedback.corrected_output
        else:
            # If correction exists, update confidence based on consistency
            existing_correction = pattern.correction_mapping[feedback.original_output]
            if existing_correction == feedback.corrected_output:
                pattern.confidence_score = min(1.0, pattern.confidence_score + 0.1)
            else:
                pattern.confidence_score = max(0.1, pattern.confidence_score - 0.1)
        
        pattern.usage_count += 1
        pattern.last_updated = datetime.now()
    
    def save_patterns(self):
        """Save patterns to database"""
        conn = sqlite3.connect(self.db.db_path)
        cursor = conn.cursor()
        
        for pattern in self.patterns.values():
            cursor.execute('''
                INSERT OR REPLACE INTO learning_patterns 
                (pattern_id, input_features, correction_mapping, confidence_score, 
                 usage_count, last_updated)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                pattern.pattern_id,
                json.dumps(pattern.input_features),
                json.dumps(pattern.correction_mapping),
                pattern.confidence_score,
                pattern.usage_count,
                pattern.last_updated.isoformat()
            ))
        
        conn.commit()
        conn.close()

class AutoCorrectEngine:
    """Engine that applies learned patterns to auto-correct outputs"""
    
    def __init__(self, learning_engine: LearningEngine):
        self.learning_engine = learning_engine
    
    def suggest_corrections(self, original_output: str, context: Dict[str, Any] = None) -> List[Tuple[str, float]]:
        """Suggest corrections based on learned patterns"""
        suggestions = []
        
        for pattern in self.learning_engine.patterns.values():
            # Check if this output matches any learned corrections
            if original_output in pattern.correction_mapping:
                correction = pattern.correction_mapping[original_output]
                confidence = pattern.confidence_score
                suggestions.append((correction, confidence))
            
            # Check for partial matches or similar patterns
            similarity = self.learning_engine.calculate_similarity(
                self.learning_engine.extract_features(original_output),
                pattern.input_features
            )
            
            if similarity > 0.6:  # Threshold for similar patterns
                for orig, corr in pattern.correction_mapping.items():
                    if orig != original_output:
                        # Apply pattern-based correction
                        suggested_correction = self.apply_pattern_correction(
                            original_output, orig, corr
                        )
                        if suggested_correction:
                            confidence = pattern.confidence_score * similarity
                            suggestions.append((suggested_correction, confidence))
        
        # Sort by confidence and return top suggestions
        suggestions.sort(key=lambda x: x[1], reverse=True)
        return suggestions[:3]  # Return top 3 suggestions
    
    def apply_pattern_correction(self, original: str, pattern_orig: str, pattern_corr: str) -> str:
        """Apply a learned correction pattern to new text"""
        # Simple pattern application - can be enhanced with more sophisticated NLP
        if pattern_orig.lower() in original.lower():
            return original.replace(pattern_orig, pattern_corr)
        return None
    
    def auto_correct(self, original_output: str, context: Dict[str, Any] = None) -> str:
        """Automatically correct output based on learned patterns"""
        suggestions = self.suggest_corrections(original_output, context)
        
        if suggestions:
            # Use the highest confidence suggestion
            best_correction, confidence = suggestions[0]
            if confidence > 0.7:  # High confidence threshold
                return best_correction
        
        return original_output

class AdaptiveAgent:
    """Main agent class that orchestrates the learning and correction system"""
    
    def __init__(self):
        self.db = FeedbackDatabase()
        self.learning_engine = LearningEngine(self.db)
        self.auto_correct_engine = AutoCorrectEngine(self.learning_engine)
    
    def generate_output(self, input_data: str, context: Dict[str, Any] = None) -> str:
        """Generate initial output (this would be your original program logic)"""
        # This is where your original program logic would go
        # For demonstration, we'll use a simple example
        return f"Generated output for: {input_data}"
    
    def process_user_correction(self, original_output: str, corrected_output: str, 
                              user_id: str, context: Dict[str, Any] = None):
        """Process user correction and update learning"""
        feedback = UserFeedback(
            original_output=original_output,
            corrected_output=corrected_output,
            user_id=user_id,
            timestamp=datetime.now(),
            context=context or {},
            feedback_type="correction"
        )
        
        # Store feedback
        self.db.store_feedback(feedback)
        
        # Process feedback for learning
        self.learning_engine.process_feedback(feedback)
        
        print(f"Learned from user correction: '{original_output}' -> '{corrected_output}'")
    
    def get_improved_output(self, input_data: str, context: Dict[str, Any] = None) -> str:
        """Get output with auto-corrections applied"""
        # Generate initial output
        initial_output = self.generate_output(input_data, context)
        
        # Apply auto-corrections
        corrected_output = self.auto_correct_engine.auto_correct(initial_output, context)
        
        return corrected_output
    
    def get_learning_stats(self) -> Dict[str, Any]:
        """Get statistics about the learning system"""
        feedback_count = len(self.db.get_feedback_history())
        pattern_count = len(self.learning_engine.patterns)
        
        return {
            "total_feedback": feedback_count,
            "learned_patterns": pattern_count,
            "patterns": [
                {
                    "id": pattern.pattern_id,
                    "confidence": pattern.confidence_score,
                    "usage_count": pattern.usage_count,
                    "corrections": len(pattern.correction_mapping)
                }
                for pattern in self.learning_engine.patterns.values()
            ]
        }

# Example usage and deployment
def main():
    """Example of how to use the adaptive agent system"""
    
    # Initialize the agent
    agent = AdaptiveAgent()
    
    # Simulate some initial interactions
    print("=== Initial Agent Output ===")
    output1 = agent.get_improved_output("Hello world")
    print(f"Output: {output1}")
    
    # Simulate user correction
    print("\n=== User Provides Correction ===")
    agent.process_user_correction(
        original_output=output1,
        corrected_output="Hello, world!",
        user_id="user1",
        context={"task_type": "greeting"}
    )
    
    # Agent learns and improves
    print("\n=== Agent Learns and Improves ===")
    output2 = agent.get_improved_output("Hi there")
    print(f"Improved Output: {output2}")
    
    # Show learning statistics
    print("\n=== Learning Statistics ===")
    stats = agent.get_learning_stats()
    print(f"Total feedback received: {stats['total_feedback']}")
    print(f"Learned patterns: {stats['learned_patterns']}")
    
    # Simulate more corrections to show learning
    print("\n=== More Learning Examples ===")
    agent.process_user_correction(
        original_output="Good morning",
        corrected_output="Good morning!",
        user_id="user1",
        context={"task_type": "greeting"}
    )
    
    # Test auto-correction
    output3 = agent.get_improved_output("Good evening")
    print(f"Auto-corrected Output: {output3}")

if __name__ == "__main__":
    main()