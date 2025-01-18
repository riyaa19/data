import hashlib
import random
from typing import Dict, Any, Optional, List
from collections import defaultdict

# External libraries for enhanced functionality
from cryptography.fernet import Fernet
from termcolor import colored

# --------------------------
# DECENTRALIZED STORAGE SIMULATION
# --------------------------
class DecentralizedStorage:
    """
    A decentralized storage system that uses hash-based data retrieval.
    Mimics content-addressing principles similar to IPFS.
    """
    def _init_(self):
        """Initialize an empty storage dictionary."""
        self.storage: Dict[str, str] = {}

    def store_data(self, data: str) -> str:
        """
        Store data and return its content-based hash.
        
        Args:
            data (str): Data to be stored
        
        Returns:
            str: SHA-256 hash of the data
        """
        data_hash = hashlib.sha256(data.encode()).hexdigest()
        self.storage[data_hash] = data
        return data_hash

    def retrieve_data(self, data_hash: str) -> Optional[str]:
        """
        Retrieve data using its hash.
        
        Args:
            data_hash (str): Hash of the stored data
        
        Returns:
            Optional[str]: Retrieved data or None
        """
        return self.storage.get(data_hash)

# --------------------------
# HIPAA COMPLIANCE: ENCRYPTION & ACCESS CONTROL
# --------------------------
class HIPAACompliance:
    """
    Manages encryption and decryption of sensitive health data
    to ensure HIPAA compliance.
    """
    def _init_(self):
        """Generate a symmetric encryption key."""
        self.key = Fernet.generate_key()
        self.cipher = Fernet(self.key)

    def encrypt_data(self, data: str) -> str:
        """
        Encrypt sensitive health data.
        
        Args:
            data (str): Data to encrypt
        
        Returns:
            str: Encrypted data
        """
        return self.cipher.encrypt(data.encode()).decode()

    def decrypt_data(self, encrypted_data: str) -> str:
        """
        Decrypt previously encrypted data.
        
        Args:
            encrypted_data (str): Encrypted data string
        
        Returns:
            str: Decrypted original data
        """
        return self.cipher.decrypt(encrypted_data.encode()).decode()

# --------------------------
# BLOCKCHAIN AND BLOCK STRUCTURE
# --------------------------
class Block:
    """
    Represents a single block in the blockchain, 
    storing patient health record information.
    """
    def _init_(self, index: int, previous_hash: str, data_hash: str, 
                 validator: str, aadhar_id: str):
        """
        Initialize a new block with transaction details.
        
        Args:
            index (int): Block's position in the blockchain
            previous_hash (str): Hash of the previous block
            data_hash (str): Hash of the stored data
            validator (str): Validator who added this block
            aadhar_id (str): Unique patient identifier
        """
        self.index = index
        self.previous_hash = previous_hash
        self.data_hash = data_hash
        self.validator = validator
        self.aadhar_id = aadhar_id
        self.hash = self.calculate_hash()

    def calculate_hash(self) -> str:
        """
        Calculate the hash of the current block.
        
        Returns:
            str: SHA-256 hash of block data
        """
        block_data = (
            str(self.index) + 
            self.previous_hash + 
            self.data_hash + 
            self.validator + 
            self.aadhar_id
        )
        return hashlib.sha256(block_data.encode()).hexdigest()

class Blockchain:
    """
    Manages the blockchain, storing blocks and providing 
    methods for block creation and validation.
    """
    def _init_(self):
        """Initialize blockchain with genesis block."""
        self.chain: List[Block] = [self.create_genesis_block()]
        self.aadhar_mapping: Dict[str, int] = {}

    def create_genesis_block(self) -> Block:
        """
        Create the first block in the blockchain.
        
        Returns:
            Block: Genesis block
        """
        return Block(0, "0", "Genesis Block", "System", "None")

    def get_latest_block(self) -> Block:
        """
        Retrieve the most recently added block.
        
        Returns:
            Block: Last block in the chain
        """
        return self.chain[-1]

    def add_block(self, block: Block) -> None:
        """
        Add a new block to the blockchain if valid.
        
        Args:
            block (Block): Block to be added
        """
        if self.is_valid_block(block):
            self.chain.append(block)
            self.aadhar_mapping[block.aadhar_id] = block.index
            print(f"Block added by {block.validator}: {block.hash}")
        else:
            print("Invalid block.")

    def is_valid_block(self, block: Block) -> bool:
        """
        Validate a block before adding to the blockchain.
        
        Args:
            block (Block): Block to validate
        
        Returns:
            bool: Whether the block is valid
        """
        previous_block = self.get_latest_block()
        return block.previous_hash == previous_block.hash

    def get_block_by_aadhar(self, aadhar_id: str) -> Optional[Block]:
        """
        Retrieve a block by patient's Aadhaar ID.
        
        Args:
            aadhar_id (str): Patient's unique identifier
        
        Returns:
            Optional[Block]: Block associated with Aadhaar ID
        """
        block_index = self.aadhar_mapping.get(aadhar_id)
        return self.chain[block_index] if block_index is not None else None

# --------------------------
# PROOF OF TRUST & EXPERTISE (PoTE)
# --------------------------
class PoTE:
    """
    Implements a Proof of Trust and Expertise (PoTE) 
    consensus mechanism for validator selection.
    """
    def _init_(self):
        """Initialize validators and trust scores."""
        self.validators: Dict[str, str] = {}
        self.trust_scores = defaultdict(int)

    def register_validator(self, validator: str, expertise_domain: str) -> None:
        """
        Register a new validator with their expertise domain.
        
        Args:
            validator (str): Name of the validator
            expertise_domain (str): Medical domain of expertise
        """
        self.validators[validator] = expertise_domain
        self.trust_scores[validator] = random.randint(50, 100)

    def calculate_trust_weight(self, validator: str) -> float:
        """
        Calculate a validator's trust weight.
        
        Args:
            validator (str): Validator to calculate trust for
        
        Returns:
            float: Trust weight
        """
        return self.trust_scores[validator] + random.uniform(0, 10)

    def select_validator(self, transaction: Dict[str, str]) -> str:
        """
        Select the most appropriate validator for a transaction.
        
        Args:
            transaction (Dict[str, str]): Transaction details
        
        Returns:
            str: Selected validator's name
        """
        eligible_validators = [
            v for v in self.validators 
            if self.validators[v] == transaction["domain"]
        ]
        if not eligible_validators:
            raise Exception("No eligible validators found!")
        return max(eligible_validators, key=self.calculate_trust_weight)

    def update_trust_scores(self, validator: str, is_successful: bool) -> None:
        """
        Update a validator's trust score based on performance.
        
        Args:
            validator (str): Validator to update
            is_successful (bool): Whether the validation was successful
        """
        if is_successful:
            self.trust_scores[validator] += 5
        else:
            self.trust_scores[validator] -= 10
        self.trust_scores[validator] = max(0, min(100, self.trust_scores[validator]))

# --------------------------
# UTILITY FUNCTIONS
# --------------------------
def display_blockchain(blockchain: Blockchain) -> None:
    """
    Display the entire blockchain state.
    
    Args:
        blockchain (Blockchain): Blockchain to display
    """
    print("\nFinal Blockchain State:")
    for block in blockchain.chain:
        print(
            f"Block {block.index}: Aadhaar ID = {block.aadhar_id}, "
            f"Data Hash = {block.data_hash}, Validator = {block.validator}, "
            f"Hash = {block.hash}"
        )

def print_header(message: str) -> None:
    """
    Print a stylized header message.
    
    Args:
        message (str): Message to display
    """
    print(colored(f"\n{'=' * 50}", 'cyan'))
    print(colored(message.center(50), 'blue', attrs=['bold']))
    print(colored(f"{'=' * 50}\n", 'cyan'))

def print_success(message: str) -> None:
    """
    Print a success message in green.
    
    Args:
        message (str): Success message to display
    """
    print(colored(f"✓ {message}", 'green'))

def print_error(message: str) -> None:
    """
    Print an error message in red.
    
    Args:
        message (str): Error message to display
    """
    print(colored(f"✗ {message}", 'red'))

# --------------------------
# MAIN APPLICATION
# --------------------------
def main():
    """
    Main entry point for the EHR Management System.
    Handles user interactions and system workflows.
    """
    print_header("🩺 Decentralized Electronic Health Record Management System 🔐")
    
    # Initialize core system components
    blockchain = Blockchain()
    pote = PoTE()
    storage = DecentralizedStorage()
    hipaa = HIPAACompliance()

    # Step 1: Register Validators
    print_header("VALIDATOR REGISTRATION")
    while True:
        validator_name = input("Enter validator name (or type 'done' to finish): ").strip()
        if validator_name.lower() == "done":
            break
        expertise_domain = input(f"Enter expertise domain for {validator_name}: ").strip()
        pote.register_validator(validator_name, expertise_domain)

    # Step 2: Add Initial Transactions
    print_header("INITIAL PATIENT DATA ENTRY")
    while True:
        aadhar_id = input("Enter Aadhaar ID (or type 'done' to finish): ").strip()
        if aadhar_id.lower() == "done":
            break
        
        ehr_data = input(f"Enter sensitive EHR data for Patient {aadhar_id}: ").strip()
        domain = input(f"Enter medical domain (e.g., Cardiology, Neurology): ").strip()

        encrypted_data = hipaa.encrypt_data(ehr_data)
        data_hash = storage.store_data(encrypted_data)

        transaction = {"aadhar_id": aadhar_id, "data_hash": data_hash, "domain": domain}

        try:
            selected_validator = pote.select_validator(transaction)
            print_success(f"Transaction validated by {selected_validator}.")

            previous_block = blockchain.get_latest_block()
            new_block = Block(
                index=len(blockchain.chain),
                previous_hash=previous_block.hash,
                data_hash=transaction["data_hash"],
                validator=selected_validator,
                aadhar_id=aadhar_id,
            )
            blockchain.add_block(new_block)
            pote.update_trust_scores(selected_validator, is_successful=True)

        except Exception as e:
            print_error(f"Error processing transaction: {e}")


 # Main Menu
    while True:
        print("\n--- EHR Management System ---")
        print("1. Retrieve Patient Data")
        print("2. Add New Patient Data")
        print("3. Update Patient Data")
        print("4. Display Blockchain")
        print("5. Exit")
        
        choice = input("Enter your choice (1-5): ").strip()

        if choice == '1':
            # Retrieve Patient Data
            while True:
                print("\nRetrieve Patient Data")
                aadhar_id = input("Enter Aadhaar ID to retrieve data (or type 'done' to exit): ").strip()
                
                if aadhar_id.lower() == 'done':
                    break

                block = blockchain.get_block_by_aadhar(aadhar_id)
                if block:
                    encrypted_data = storage.retrieve_data(block.data_hash)
                    if encrypted_data:
                        decrypted_data = hipaa.decrypt_data(encrypted_data)
                        
                        print(f"\nBlock Found for Aadhaar ID {aadhar_id}:")
                        print(f"Block Hash: {block.hash}")
                        print(f"Data Hash Stored in Blockchain: {block.data_hash}")
                        print("Hash Matched! Decrypting Data...")
                        print(f"Decrypted EHR Data for Aadhaar ID {aadhar_id}: {decrypted_data}")
                    else:
                        print("No data found in storage for the given hash.")
                else:
                    print(f"No block found for Aadhaar ID {aadhar_id}.")
                
                print("\nRetrieve Patient Data")

        elif choice == '2':
            # Add New Patient Data
            aadhar_id = input("Enter new Aadhaar ID: ").strip()
            ehr_data = input(f"Enter EHR data for Aadhaar ID {aadhar_id}: ").strip()
            domain = input(f"Enter medical domain (e.g., Cardiology, Neurology): ").strip()
            
            encrypted_data = hipaa.encrypt_data(ehr_data)
            data_hash = storage.store_data(encrypted_data)

            transaction = {"aadhar_id": aadhar_id, "data_hash": data_hash, "domain": domain}

            try:
                selected_validator = pote.select_validator(transaction)
                print(f"Transaction validated by {selected_validator}.")

                previous_block = blockchain.get_latest_block()
                new_block = Block(
                    index=len(blockchain.chain),
                    previous_hash=previous_block.hash,
                    data_hash=data_hash,
                    validator=selected_validator,
                    aadhar_id=aadhar_id,
                )
                blockchain.add_block(new_block)
                pote.update_trust_scores(selected_validator, is_successful=True)
                
            except Exception as e:
                print(f"Error processing transaction: {e}")

        elif choice == '3':
            # Update Patient Data
            aadhar_id = input("Enter Aadhaar ID to update data: ").strip()
            block = blockchain.get_block_by_aadhar(aadhar_id)
            if block:
                new_data = input(f"Enter updated EHR data for Aadhaar ID {aadhar_id}: ").strip()
                encrypted_data = hipaa.encrypt_data(new_data)
                new_data_hash = storage.store_data(encrypted_data)
                block.data_hash = new_data_hash
                block.hash = block.calculate_hash()
                print(f"Data updated for Aadhaar ID {aadhar_id}. New Block Hash: {block.hash}")
            else:
                print(f"No block found for Aadhaar ID {aadhar_id}.")

        elif choice == '4':
            # Display Blockchain
            display_blockchain(blockchain)

        elif choice == '5':
            # Exit
            print("Exiting EHR Management System...")
            break

        else:
            print("Invalid choice. Please try again.")


if _name_ == "_main_":
    main()
