import PySimpleGUI as sg
import random
from collections import defaultdict

class PokerPersonality:
    def __init__(self, name=None):
        self.name = name or self.generate_ai_name()
        self.traits = self.generate_personality_traits()
        self.playing_style = self.generate_playing_style()
        self.tell_system = self.generate_tells()
        self.mood_system = MoodSystem()
        self.learning_rate = random.uniform(0.1, 0.3)
        self.tilt_factor = 0.0  # Increases with bad beats
        
    def generate_ai_name(self):
        """Generate a poker-themed name for the AI"""
        prefixes = ['Lucky', 'Wild', 'Cool', 'Smooth', 'Sharp', 'Quick']
        suffixes = ['Ace', 'King', 'Queen', 'Jack', 'River', 'Flop']
        return f"{random.choice(prefixes)} {random.choice(suffixes)}"
        
    def generate_personality_traits(self):
        """Generate random personality traits"""
        return {
            'aggression': random.uniform(0.3, 0.8),
            'patience': random.uniform(0.4, 0.9),
            'discipline': random.uniform(0.4, 0.9),
            'adaptability': random.uniform(0.3, 0.8),
            'risk_tolerance': random.uniform(0.2, 0.8),
            
            'psychological': {
                'tilt_resistance': random.uniform(0.3, 0.9),
                'pressure_handling': random.uniform(0.4, 0.9),
                'bluff_detection': random.uniform(0.3, 0.8),
                'table_presence': random.uniform(0.4, 0.9),
                'emotional_control': random.uniform(0.5, 0.9)
            }
        }

class MoodSystem:
    def __init__(self):
        self.base_mood = 0.5  # Neutral mood
        self.current_mood = self.base_mood
        self.mood_history = []
        self.tilt_threshold = 0.8
        self.recovery_rate = 0.1
        
    def update_mood(self, event_impact):
        """Update mood based on game events"""
        self.current_mood = max(0, min(1, self.current_mood + event_impact))
        self.mood_history.append(self.current_mood)
        
        if len(self.mood_history) > 10:
            self.mood_history.pop(0)
            
    def is_tilted(self):
        """Check if AI is on tilt"""
        return self.current_mood > self.tilt_threshold
        
    def recover(self):
        """Gradually recover mood towards base"""
        if self.current_mood != self.base_mood:
            direction = 1 if self.current_mood < self.base_mood else -1
            self.current_mood += direction * self.recovery_rate
            self.current_mood = max(0, min(1, self.current_mood))

class PokerGame:
    def __init__(self):
        self.window = None
        self.chips = 1000
        self.current_bet = 0
        self.pot = 0
        self.deck = self.create_deck()
        self.player_hand = []
        self.dealer_hand = []
        self.game_phase = 'bet'  # bet, flop, turn, river
        self.community_cards = []
        self.dealer_ai = AdvancedDealerAI()  # Initialize the AI
        self.stats = AdvancedStats()
        self.last_action = None
        
    def create_deck(self):
        """Create a new deck of cards"""
        suits = ['♠', '♣', '♥', '♦']
        ranks = ['2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K', 'A']
        deck = [(rank, suit) for suit in suits for rank in ranks]
        random.shuffle(deck)
        return deck
        
    def deal_cards(self):
        """Deal initial cards"""
        self.player_hand = [self.deck.pop(), self.deck.pop()]
        self.dealer_hand = [self.deck.pop(), self.deck.pop()]
        
    def show(self):
        """Display the poker game window"""
        if self.window is not None:
            self.window.close()
            
        layout = [
            [sg.Text(f'Chips: {self.chips}', key='-CHIPS-')],
            [sg.Text('Your Hand:', key='-HAND-LABEL-')],
            [sg.Text('', key='-PLAYER-HAND-')],
            [sg.Text('Community Cards:', key='-COMMUNITY-LABEL-')],
            [sg.Text('', key='-COMMUNITY-CARDS-')],
            [sg.Text('Dealer:', key='-DEALER-LABEL-')],
            [sg.Text('', key='-DEALER-MOOD-')],  # Show dealer's mood
            [sg.Text(f'Pot: {self.pot}', key='-POT-')],
            [
                sg.Button('Deal', key='-DEAL-'),
                sg.Button('Check', key='-CHECK-', disabled=True),
                sg.Button('Call', key='-CALL-', disabled=True),
                sg.Button('Raise', key='-RAISE-', disabled=True),
                sg.Button('Fold', key='-FOLD-', disabled=True),
                sg.Button('Exit', key='-EXIT-')
            ],
            [sg.Text('', key='-STATUS-')]  # Status messages
        ]
        
        self.window = sg.Window('Poker Game', layout, finalize=True)
        self.run_game_loop()
        
    def run_game_loop(self):
        """Main game loop"""
        while True:
            event, values = self.window.read()
            
            if event in (None, '-EXIT-', sg.WIN_CLOSED):
                break
                
            if event == '-DEAL-':
                self.start_new_hand()
                
            elif event == '-CHECK-':
                self.handle_player_action('check', 0)
                
            elif event == '-CALL-':
                self.handle_player_action('call', self.current_bet)
                
            elif event == '-RAISE-':
                amount = self.get_raise_amount()
                if amount:
                    self.handle_player_action('raise', amount)
                    
            elif event == '-FOLD-':
                self.handle_player_action('fold', 0)
                
        self.window.close()
        self.window = None
        
    def handle_player_action(self, action, amount):
        """Handle player actions and get AI response"""
        self.last_action = action
        game_state = self.get_game_state()
        
        # Update pot and chips based on player action
        if action in ['call', 'raise']:
            self.pot += amount
            self.chips -= amount
            
        # Let AI analyze player's action
        self.dealer_ai.analyze_player_patterns(action, amount)
        
        # Get AI's response
        ai_action, ai_amount = self.dealer_ai.make_decision(game_state, action)
        self.handle_ai_response(ai_action, ai_amount)
        
        # Update display
        self.update_display()
        self.update_game_phase()
        
    def handle_ai_response(self, action, amount):
        """Handle AI's response to player action"""
        status_msg = f"Dealer {action}s"
        if amount > 0:
            status_msg += f" with {amount} chips"
            
        self.window['-STATUS-'].update(status_msg)
        
        if action == 'raise':
            self.pot += amount
            self.current_bet = amount
            self.enable_response_buttons()
        elif action == 'call':
            self.pot += amount
            self.advance_game_phase()
        elif action == 'fold':
            self.handle_win(self.pot)
            
    def get_game_state(self):
        """Get current game state for AI decision making"""
        return {
            'pot': self.pot,
            'current_bet': self.current_bet,
            'dealer_hand': self.dealer_hand,
            'community_cards': self.community_cards,
            'game_phase': self.game_phase,
            'player_chips': self.chips,
            'last_action': self.last_action
        }
        
    def advance_game_phase(self):
        """Advance to next phase of the game"""
        if self.game_phase == 'bet':
            self.deal_flop()
        elif self.game_phase == 'flop':
            self.deal_turn()
        elif self.game_phase == 'turn':
            self.deal_river()
        elif self.game_phase == 'river':
            self.handle_showdown()
            
    def update_display(self):
        """Update the display with current game state"""
        # Update player's hand
        hand_str = ' '.join([f'{rank}{suit}' for rank, suit in self.player_hand])
        self.window['-PLAYER-HAND-'].update(hand_str)
        
        # Update community cards
        if self.community_cards:
            community_str = ' '.join([f'{rank}{suit}' for rank, suit in self.community_cards])
            self.window['-COMMUNITY-CARDS-'].update(community_str)
            
        # Update chips and pot
        self.window['-CHIPS-'].update(f'Chips: {self.chips}')
        self.window['-POT-'].update(f'Pot: {self.pot}')
        
        # Update dealer mood
        mood = "🤔" if self.dealer_ai.playing_style.get('confidence', 0.5) < 0.3 else "😎"
        self.window['-DEALER-MOOD-'].update(f'Mood: {mood}')
        
    def enable_response_buttons(self):
        """Enable response buttons after AI's response"""
        self.window['-CALL-'].update(disabled=False)
        self.window['-RAISE-'].update(disabled=False)
        self.window['-FOLD-'].update(disabled=False)
        self.window['-CHECK-'].update(disabled=True)
        
    def start_new_hand(self):
        """Start a new hand"""
        self.deck = self.create_deck()
        self.deal_cards()
        self.update_display()
        self.enable_game_buttons()
        
    def enable_game_buttons(self):
        """Enable game buttons after dealing"""
        self.window['-CHECK-'].update(disabled=False)
        self.window['-CALL-'].update(disabled=False)
        self.window['-RAISE-'].update(disabled=False)
        self.window['-FOLD-'].update(disabled=False)
        self.window['-DEAL-'].update(disabled=True) 

class AdvancedStats:
    def __init__(self):
        self.hands_played = 0
        self.hands_won = 0
        self.chips_won = 0
        self.biggest_pot = 0
        self.biggest_bluff = 0
        self.hand_history = []
        self.position_stats = defaultdict(lambda: {'played': 0, 'won': 0})
        self.hand_types = defaultdict(int)
        
    def update_stats(self, game_result):
        """Update statistics after each hand"""
        self.hands_played += 1
        if game_result['won']:
            self.hands_won += 1
            self.chips_won += game_result['amount']
            
        if game_result['pot'] > self.biggest_pot:
            self.biggest_pot = game_result['pot']
            
        self.hand_history.append(game_result)
        self.position_stats[game_result['position']]['played'] += 1
        if game_result['won']:
            self.position_stats[game_result['position']]['won'] += 1
            
        self.hand_types[game_result['hand_type']] += 1

class AIPersonalityTraits:
    def __init__(self):
        self.playing_styles = {
            'aggressive': {
                'risk_tolerance': 0.8,
                'bluff_frequency': 0.7,
                'aggression': 0.8,
                'bet_sizing': 0.7
            },
            'conservative': {
                'risk_tolerance': 0.3,
                'bluff_frequency': 0.2,
                'aggression': 0.2,
                'bet_sizing': 0.4
            },
            'balanced': {
                'risk_tolerance': 0.5,
                'bluff_frequency': 0.4,
                'aggression': 0.5,
                'bet_sizing': 0.5
            },
            'confident': {
                'aggression': 0.1,
                'risk_tolerance': 0.1,
                'bluff_frequency': -0.1
            },
            'cautious': {
                'aggression': -0.2,
                'risk_tolerance': -0.2,
                'bluff_frequency': -0.2
            },
            'neutral': {
                'aggression': 0,
                'risk_tolerance': 0,
                'bluff_frequency': 0
            }
        }

class AdvancedDealerAI:
    def __init__(self):
        self.playing_style = self.generate_ai_personality()
        self.hand_history = []
        self.player_patterns = {
            'betting_frequency': [],
            'average_bet_size': [],
            'bluff_suspected': [],
            'position_aggression': {'early': 0, 'middle': 0, 'late': 0}
        }
        self.session_stats = {'hands_won': 0, 'chips_won': 0, 'biggest_bluff': 0}
        self.mood_adjustments = {
            'winning_streak': {'confidence': 0.1, 'aggression': 0.05},
            'losing_streak': {'confidence': -0.1, 'caution': 0.1},
            'big_win': {'confidence': 0.15, 'aggression': 0.1},
            'big_loss': {'confidence': -0.15, 'caution': 0.15}
        }
        
    def generate_ai_personality(self):
        """Generate a random personality for the AI dealer"""
        traits = AIPersonalityTraits()
        style_keys = list(traits.playing_styles.keys())
        base_style = random.choice(style_keys)
        return traits.playing_styles[base_style]
        
    def analyze_player_patterns(self, action, bet_size=0):
        """Analyze and store player betting patterns"""
        self.player_patterns['betting_frequency'].append(action == 'bet')
        if bet_size > 0:
            self.player_patterns['average_bet_size'].append(bet_size)
            
    def adjust_strategy(self, game_state):
        """Adjust AI strategy based on game state and player patterns"""
        if len(self.player_patterns['betting_frequency']) > 10:
            player_aggression = sum(self.player_patterns['betting_frequency'][-10:]) / 10
            if player_aggression > 0.7:
                self.playing_style['caution'] = min(1, self.playing_style.get('caution', 0) + 0.1)
            elif player_aggression < 0.3:
                self.playing_style['aggression'] = min(1, self.playing_style.get('aggression', 0) + 0.1)
                
    def update_mood(self, result):
        """Update AI's mood based on game results"""
        if result['type'] == 'win':
            self.apply_mood_adjustment('winning_streak')
        elif result['type'] == 'loss':
            self.apply_mood_adjustment('losing_streak')
            
        if result.get('amount', 0) > 1000:
            if result['type'] == 'win':
                self.apply_mood_adjustment('big_win')
            else:
                self.apply_mood_adjustment('big_loss')
                
    def apply_mood_adjustment(self, mood_type):
        """Apply mood-based adjustments to AI behavior"""
        if mood_type in self.mood_adjustments:
            adjustments = self.mood_adjustments[mood_type]
            for trait, adjustment in adjustments.items():
                current = self.playing_style.get(trait, 0.5)
                self.playing_style[trait] = max(0, min(1, current + adjustment))

    def make_decision(self, game_state, player_action):
        """Make a decision based on game state and adjusted strategy"""
        hand_strength = self.evaluate_hand_strength(game_state['dealer_hand'], game_state['community_cards'])
        pot_odds = self.calculate_pot_odds(game_state['pot'], game_state['current_bet'])
        
        # Factor in personality and mood
        confidence = self.playing_style.get('confidence', 0.5)
        aggression = self.playing_style.get('aggression', 0.5)
        caution = self.playing_style.get('caution', 0.5)
        
        # Calculate base action threshold
        action_threshold = (hand_strength * confidence) - (pot_odds * caution) + (aggression * 0.2)
        
        if action_threshold > 0.7:
            return 'raise', self.calculate_bet_size(game_state)
        elif action_threshold > 0.4:
            return 'call', game_state['current_bet']
        else:
            return 'fold', 0
            
    def evaluate_hand_strength(self, hand, community_cards):
        """Evaluate the strength of the current hand"""
        # Implement hand strength evaluation logic
        return random.random()  # Placeholder
        
    def calculate_pot_odds(self, pot, bet):
        """Calculate pot odds for decision making"""
        if pot + bet == 0:
            return 0
        return bet / (pot + bet)
        
    def calculate_bet_size(self, game_state):
        """Calculate appropriate bet size based on strategy and game state"""
        base_bet = game_state['pot'] * 0.5
        aggression_factor = self.playing_style.get('aggression', 0.5)
        return int(base_bet * (1 + aggression_factor))

# Add any other poker-related classes and functions here 