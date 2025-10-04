"""
Web UI for Adaptive Agent System
Provides a user interface for interacting with the learning agent
"""

from flask import Flask, render_template, request, jsonify, session
import json
from adaptive_agent_system import AdaptiveAgent
import uuid
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Change this in production

# Initialize the adaptive agent
agent = AdaptiveAgent()

@app.route('/')
def index():
    """Main interface page"""
    return render_template('index.html')

@app.route('/api/generate', methods=['POST'])
def generate_output():
    """Generate output from the agent"""
    try:
        data = request.get_json()
        input_text = data.get('input_text', '')
        context = data.get('context', {})
        
        # Generate output
        output = agent.get_improved_output(input_text, context)
        
        # Store in session for potential correction
        session['last_input'] = input_text
        session['last_output'] = output
        session['last_context'] = context
        
        return jsonify({
            'success': True,
            'output': output,
            'timestamp': datetime.now().isoformat()
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/correct', methods=['POST'])
def correct_output():
    """Process user correction"""
    try:
        data = request.get_json()
        corrected_output = data.get('corrected_output', '')
        user_id = session.get('user_id', str(uuid.uuid4()))
        
        # Get the last output from session
        original_output = session.get('last_output', '')
        context = session.get('last_context', {})
        
        if not original_output:
            return jsonify({
                'success': False,
                'error': 'No previous output to correct'
            }), 400
        
        # Process the correction
        agent.process_user_correction(
            original_output=original_output,
            corrected_output=corrected_output,
            user_id=user_id,
            context=context
        )
        
        # Update session
        session['user_id'] = user_id
        
        return jsonify({
            'success': True,
            'message': 'Correction learned successfully',
            'timestamp': datetime.now().isoformat()
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/suggestions', methods=['POST'])
def get_suggestions():
    """Get auto-correction suggestions"""
    try:
        data = request.get_json()
        output_text = data.get('output_text', '')
        context = data.get('context', {})
        
        suggestions = agent.auto_correct_engine.suggest_corrections(output_text, context)
        
        return jsonify({
            'success': True,
            'suggestions': [
                {'text': suggestion, 'confidence': confidence}
                for suggestion, confidence in suggestions
            ]
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/learn-from-edit', methods=['POST'])
def learn_from_edit():
    """Learn from pencil edit corrections"""
    try:
        data = request.get_json()
        original_output = data.get('original_output', '')
        corrected_output = data.get('corrected_output', '')
        context = data.get('learning_context', {})
        
        user_id = session.get('user_id', str(uuid.uuid4()))
        
        if not original_output or not corrected_output:
            return jsonify({
                'success': False,
                'error': 'Both original and corrected output required'
            }), 400
        
        # Add pencil edit context
        context.update({
            'feedback_type': 'pencil_edit',
            'edit_method': 'inline_editing'
        })
        
        # Process the correction
        agent.process_user_correction(
            original_output=original_output,
            corrected_output=corrected_output,
            user_id=user_id,
            context=context
        )
        
        # Update session
        session['user_id'] = user_id
        
        # Get learning info for response
        stats = agent.get_learning_stats()
        latest_pattern = None
        if stats['patterns']:
            latest_pattern = max(stats['patterns'], 
                               key=lambda p: p.get('usage_count', 0))
        
        return jsonify({
            'success': True,
            'message': 'Learning completed successfully',
            'confidence': latest_pattern['confidence'] if latest_pattern else 0.5,
            'pattern_id': latest_pattern['id'] if latest_pattern else 'new',
            'timestamp': datetime.now().isoformat()
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/stats')
def get_stats():
    """Get learning statistics"""
    try:
        stats = agent.get_learning_stats()
        return jsonify({
            'success': True,
            'stats': stats
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)